import argparse
import json
import shutil
import tempfile
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook


def get_link(cell):
    link = cell.hyperlink.target if cell.hyperlink else cell.value
    if isinstance(link, str):
        return link.strip()
    return None


def download_pdf(url, destination):
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        raise FileExistsError(
            f"PDF-filen findes allerede og overskrives ikke: {destination}"
        )

    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=destination.parent, prefix="download_", suffix=".tmp", delete=False
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
            request = urllib.request.Request(
                url, headers={"User-Agent": "Mozilla/5.0"}
            )
            with urllib.request.urlopen(request, timeout=30) as response:
                shutil.copyfileobj(response, temporary_file)

        temporary_path.replace(destination)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def write_log(log_path, entries):
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=log_path.parent,
            prefix="log_",
            suffix=".tmp",
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
            json.dump(entries, temporary_file, ensure_ascii=False, indent=2)
            temporary_file.write("\n")

        temporary_path.replace(log_path)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def copy_pdf_file(source, destination):
    temporary_path = None
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        with source.open("rb") as source_file:
            if source_file.read(5) != b"%PDF-":
                raise ValueError(f"Filen er ikke en PDF: {source}")
            source_file.seek(0)
            with tempfile.NamedTemporaryFile(
                mode="wb",
                dir=destination.parent,
                prefix="pdf_backup_",
                suffix=".tmp",
                delete=False,
            ) as temporary_file:
                temporary_path = Path(temporary_file.name)
                shutil.copyfileobj(source_file, temporary_file)
        temporary_path.replace(destination)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def record_backup_error(error_log_path, error):
    try:
        if error_log_path.exists():
            with error_log_path.open("r", encoding="utf-8") as error_file:
                errors = json.load(error_file)
            if not isinstance(errors, list) or any(
                not isinstance(entry, dict) for entry in errors
            ):
                raise ValueError(
                    f"{error_log_path.name} skal indeholde en liste af fejlobjekter."
                )
        else:
            errors = []

        errors.append(
            {
                "error_type": f"{type(error).__name__}: {error}",
                "date": datetime.now().astimezone().isoformat(timespec="seconds"),
            }
        )
        write_log(error_log_path, errors)
    except (OSError, json.JSONDecodeError, ValueError, TypeError) as logging_error:
        print(
            f"FEJL: Kunne ikke logge backup-fejl i {error_log_path}: "
            f"{logging_error}"
        )


def update_backup(backup_path, backup_folder, backup_error_log_path, entries):
    backup_entries = []
    try:
        backup_folder.mkdir(parents=True, exist_ok=True)
    except (OSError, TypeError, ValueError) as error:
        record_backup_error(backup_error_log_path, error)
        print(f"FEJL: Kunne ikke oprette PDF-backupmappen {backup_folder}: {error}")

    for original_entry in entries:
        entry = dict(original_entry)
        source_value = entry.get("file")
        if not isinstance(source_value, str) or not source_value.strip():
            entry["backup_status"] = "failed"
            entry["backup_reason"] = "Logposten mangler en gyldig PDF-filsti."
            backup_entries.append(entry)
            error = ValueError(entry["backup_reason"])
            record_backup_error(backup_error_log_path, error)
            print(f"FEJL: Kan ikke sikkerhedskopiere post {entry!r}: {error}")
            continue

        source = Path(source_value)
        if source.suffix.casefold() != ".pdf":
            entry["backup_status"] = "skipped"
            entry["backup_reason"] = "Filen har ikke .pdf-endelse og ignoreres."
            entry.pop("backup_file", None)
            backup_entries.append(entry)
            print(f"Ignorerer ikke-PDF-fil i backup: {source}")
            continue

        backup_destination = backup_folder / source.name
        try:
            copy_pdf_file(source, backup_destination)
        except ValueError as error:
            entry["backup_status"] = "skipped"
            entry["backup_reason"] = str(error)
            entry.pop("backup_file", None)
            print(f"Ignorerer ikke-PDF-fil i backup: {source} ({error})")
        except OSError as error:
            entry["backup_status"] = "failed"
            entry["backup_reason"] = f"{type(error).__name__}: {error}"
            record_backup_error(backup_error_log_path, error)
            print(f"FEJL: Kunne ikke sikkerhedskopiere {source}: {error}")
        else:
            entry["backup_status"] = "saved"
            entry["backup_file"] = str(backup_destination)
            entry.pop("backup_reason", None)
            print(f"PDF-backup opdateret: {backup_destination}")
        backup_entries.append(entry)

    try:
        write_log(backup_path, backup_entries)
    except (OSError, TypeError, ValueError) as error:
        record_backup_error(backup_error_log_path, error)
        print(f"FEJL: Kunne ikke opdatere backupfilen {backup_path}: {error}")


def get_pdf_filename(cell, br_number):
    value = str(cell.value or "").strip()
    filename = Path(value).name if value else f"{br_number}.pdf"
    if filename in {"", ".", ".."}:
        raise ValueError(f"Ugyldigt PDF-filnavn for {br_number}: {value!r}")
    if Path(filename).suffix.casefold() != ".pdf":
        filename = f"{Path(filename).stem}.pdf"
    return filename


def download_new_pdfs(
    worksheet,
    log_path,
    error_log_path,
    download_folder,
    backup_path,
    backup_folder,
    backup_error_log_path,
):
    br_number_column = None
    for cell in worksheet[1]:
        header = str(cell.value or "").strip().casefold().replace("_", "-")
        if header in {"br-number", "br number"}:
            br_number_column = cell.column
            break

    if br_number_column is None:
        raise ValueError("Kunne ikke finde kolonnen 'BR-Number' i første række.")

    if log_path.exists():
        with log_path.open("r", encoding="utf-8") as log_file:
            entries = json.load(log_file)
        if not isinstance(entries, list):
            raise ValueError("JSON-logfilen skal indeholde en liste af poster.")
        if any(not isinstance(entry, dict) for entry in entries):
            raise ValueError("Alle poster i JSON-logfilen skal være objekter.")
    else:
        entries = []

    update_backup(
        backup_path, backup_folder, backup_error_log_path, entries
    )

    if error_log_path.exists():
        with error_log_path.open("r", encoding="utf-8") as error_log_file:
            error_entries = json.load(error_log_file)
        if not isinstance(error_entries, list):
            raise ValueError("error_log.json skal indeholde en liste af fejl.")
        if any(not isinstance(entry, dict) for entry in error_entries):
            raise ValueError("Alle fejl i error_log.json skal være JSON-objekter.")
    else:
        error_entries = []
        write_log(error_log_path, error_entries)

    known_br_numbers = {
        entry["br_number"].strip().casefold()
        for entry in entries
        if isinstance(entry.get("br_number"), str) and entry["br_number"].strip()
    }
    known_links = {
        entry["link"].strip()
        for entry in entries
        if isinstance(entry.get("link"), str) and entry["link"].strip()
    }

    # Recognize old log entries without changing them.
    for entry in entries:
        if "br_number" in entry:
            continue
        row_number = entry.get("row")
        old_link = entry.get("link")
        if type(row_number) is int and 2 <= row_number <= worksheet.max_row:
            source_br_number = worksheet.cell(row_number, br_number_column).value
            source_link = get_link(worksheet.cell(row_number, 4))
            if (
                source_br_number is not None
                and isinstance(old_link, str)
                and source_link == old_link.strip()
            ):
                known_br_numbers.add(str(source_br_number).strip().casefold())

    downloaded_count = 0
    skipped_count = 0

    def record_download_error(br_number, link, error_type):
        error_entries.append(
            {
                "br_number": br_number,
                "link": link,
                "error_type": error_type,
                "date": datetime.now().astimezone().isoformat(timespec="seconds"),
            }
        )
        write_log(error_log_path, error_entries)
        print(f"FEJL for {br_number}: {error_type} - {link}")

    def try_download(br_number, link, destination):
        try:
            download_pdf(link, destination)
        except urllib.error.HTTPError as error:
            error_type = f"HTTP {error.code}: {error.reason}"
            record_download_error(br_number, link, error_type)
            return error.code
        except (urllib.error.URLError, TimeoutError, OSError) as error:
            error_type = f"{type(error).__name__}: {error}"
            record_download_error(br_number, link, error_type)
            return None
        return 200

    for row_number in range(2, worksheet.max_row + 1):
        br_number = worksheet.cell(row_number, br_number_column).value
        link = get_link(worksheet.cell(row_number, 4))
        if br_number is None or not link:
            continue
        if not link.lower().startswith(("http://", "https://")):
            continue

        br_number = str(br_number).strip()
        normalized_br_number = br_number.casefold()
        if normalized_br_number in known_br_numbers:
            print(f"Springer {br_number} over: BR-nummeret findes allerede i loggen.")
            skipped_count += 1
            continue
        if link in known_links:
            print(
                f"Springer {br_number} over: linket findes allerede i JSON-loggen."
            )
            skipped_count += 1
            continue

        filename = get_pdf_filename(worksheet.cell(row_number, 3), br_number)
        destination = download_folder / filename
        working_link = link
        status = try_download(br_number, working_link, destination)

        if status != 200:
            reserve_link = get_link(worksheet.cell(row_number, 5))
            if (
                not reserve_link
                or not reserve_link.lower().startswith(("http://", "https://"))
            ):
                print(
                    f"Ingen gyldig reserve-link i kolonne E for {br_number}; "
                    "går videre til næste række."
                )
                skipped_count += 1
                continue
            if reserve_link == link:
                print(
                    f"Reserve-linket er det samme som linket med 404 for {br_number}; "
                    "går videre til næste række."
                )
                skipped_count += 1
                continue
            if reserve_link in known_links:
                print(
                    f"Springer reserve-linket over for {br_number}: "
                    "linket findes allerede i JSON-loggen."
                )
                skipped_count += 1
                continue

            print(f"Prøver reserve-link i kolonne E for {br_number}.")
            status = try_download(br_number, reserve_link, destination)
            if status != 200:
                print(
                    f"Kunne ikke hente PDF for {br_number} fra reserve-linket; "
                    "går videre til næste række."
                )
                skipped_count += 1
                continue
            working_link = reserve_link
        elif status != 200:
            print(
                f"Kunne ikke hente PDF for {br_number}; "
                "går videre til næste række."
            )
            skipped_count += 1
            continue

        entry = {
            "br_number": br_number,
            "row": row_number,
            "link": working_link,
            "file": str(destination),
        }
        entries.append(entry)
        known_br_numbers.add(normalized_br_number)
        known_links.add(working_link)
        write_log(log_path, entries)
        update_backup(
            backup_path, backup_folder, backup_error_log_path, entries
        )
        downloaded_count += 1
        print(f"Hentede {br_number} fra {working_link}: {destination}")

    return downloaded_count, skipped_count


def reset_from_backup(
    backup_path,
    log_path,
    download_folder,
    backup_folder,
    backup_error_log_path,
    reset_log_path,
):
    try:
        with backup_path.open("r", encoding="utf-8") as backup_file:
            entries = json.load(backup_file)
        if not isinstance(entries, list) or any(
            not isinstance(entry, dict) for entry in entries
        ):
            raise ValueError("Backupfilen skal indeholde en liste af poster.")
    except (OSError, json.JSONDecodeError, ValueError) as error:
        record_backup_error(backup_error_log_path, error)
        raise

    log_entries = [
        {key: value for key, value in entry.items() if not key.startswith("backup_")}
        for entry in entries
    ]
    write_log(log_path, log_entries)
    download_folder.mkdir(parents=True, exist_ok=True)

    if reset_log_path.exists():
        with reset_log_path.open("r", encoding="utf-8") as reset_log_file:
            reset_entries = json.load(reset_log_file)
        if not isinstance(reset_entries, list) or any(
            not isinstance(entry, dict) for entry in reset_entries
        ):
            raise ValueError("reset_log.json skal indeholde en liste af poster.")
    else:
        reset_entries = []

    reset_date = datetime.now().astimezone().isoformat(timespec="seconds")
    reset_run = {
        "date": reset_date,
        "backup_file": str(backup_path),
        "backup_entry_count": len(entries),
        "results": [],
    }
    print(f"Backupfilen indeholder {len(entries)} post(er).")

    restored_count = 0
    for entry in entries:
        br_number = entry.get("br_number")
        link = entry.get("link")
        file_value = entry.get("file")
        result = {
            "br_number": br_number,
            "link": link,
            "file": file_value,
            "date": reset_date,
        }
        if not isinstance(file_value, str) or not file_value.strip():
            error = ValueError(f"Ugyldig post i backupfilen: {entry!r}")
            result.update(status="invalid", reason=str(error))
            reset_run["results"].append(result)
            print(f"Springer ugyldig backup-post over: {error}")
            continue

        filename = Path(file_value).name
        if filename in {"", ".", ".."}:
            error = ValueError(f"Ugyldigt filnavn i backupfilen: {file_value}")
            result.update(status="invalid", reason=str(error))
            reset_run["results"].append(result)
            print(f"Springer {br_number or 'ukendt fil'} over: {error}")
            continue
        destination = download_folder / filename
        backup_file_value = entry.get("backup_file", filename)
        if not isinstance(backup_file_value, str) or not backup_file_value.strip():
            backup_file_value = filename
        backup_filename = Path(backup_file_value).name
        backup_source = backup_folder / backup_filename
        if not isinstance(br_number, str) or not br_number.strip():
            br_number = Path(filename).stem
        if entry.get("backup_status") == "skipped":
            reason = entry.get(
                "backup_reason", "Filen blev fravalgt under backup."
            )
            result.update(status="skipped", reason=reason)
            reset_run["results"].append(result)
            print(f"Ignorerer {br_number}: {reason}")
            continue
        if not backup_source.is_file():
            reason = f"PDF-kopien mangler i backupmappen: {backup_source}"
            result.update(status="not_backed_up", reason=reason)
            reset_run["results"].append(result)
            record_backup_error(backup_error_log_path, FileNotFoundError(reason))
            print(f"Springer {br_number} over: {reason}")
            continue

        if destination.is_file():
            try:
                with destination.open("rb") as destination_file:
                    existing_is_pdf = destination_file.read(5) == b"%PDF-"
            except OSError as error:
                existing_is_pdf = False
                print(f"Kunne ikke kontrollere {destination}: {error}")
            if existing_is_pdf:
                result.update(
                    status="already_exists",
                    reason="En gyldig PDF findes allerede; den overskrives ikke.",
                )
                reset_run["results"].append(result)
                print(
                    f"Springer {br_number} over: gyldig PDF findes allerede "
                    f"({destination})."
                )
                continue

        try:
            copy_pdf_file(backup_source, destination)
        except (OSError, ValueError) as error:
            result.update(
                status="failed",
                reason=f"{type(error).__name__}: {error}",
            )
            reset_run["results"].append(result)
            record_backup_error(backup_error_log_path, error)
            print(f"FEJL ved gendannelse af {br_number} fra PDF-backup: {error}")
            continue

        restored_count += 1
        result.update(
            status="restored",
            reason=f"Gendannet fra den lokale PDF-backup: {backup_source}",
        )
        reset_run["results"].append(result)
        print(f"Gendannede {br_number} fra PDF-backup: {destination}")

    reset_entries.append(reset_run)
    try:
        write_log(reset_log_path, reset_entries)
    except (OSError, TypeError, ValueError) as error:
        print(f"FEJL: Kunne ikke skrive reset-resultater til {reset_log_path}: {error}")

    print(f"Reset færdig: gendannede {restored_count} PDF-fil(er).")


def main():
    parser = argparse.ArgumentParser(description="Download og gendan PDF-filer.")
    parser.add_argument(
        "--reset",
        action="store_true",
        help="Gendan PDF-linksloggen og manglende PDF-filer fra backup.",
    )
    args = parser.parse_args()

    folder = Path(__file__).resolve().parent
    excel_path = folder / "BR_Numbers_List.xlsx"
    log_path = folder / "PDF_links_log.json"
    error_log_path = folder / "error_log.json"
    backup_path = folder / "PDF_links_backup.json"
    backup_folder = folder / "PDF_backup"
    backup_error_log_path = folder / "error_backup.json"
    reset_log_path = folder / "reset_log.json"
    download_folder = folder / "PDFs"

    if args.reset:
        reset_from_backup(
            backup_path,
            log_path,
            download_folder,
            backup_folder,
            backup_error_log_path,
            reset_log_path,
        )
        return

    workbook = load_workbook(excel_path, read_only=False, data_only=True)
    try:
        downloaded, skipped = download_new_pdfs(
            workbook.active,
            log_path,
            error_log_path,
            download_folder,
            backup_path,
            backup_folder,
            backup_error_log_path,
        )
    finally:
        workbook.close()

    print(f"Færdig: hentede {downloaded} PDF-fil(er), sprang {skipped} over.")
    print(f"Logfil: {log_path}")
    print(f"Fejllog: {error_log_path}")
    print(f"Backupfil: {backup_path}")


if __name__ == "__main__":
    main()