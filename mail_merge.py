"""Create personalised Tea with Nessie invitations from Excel and Word."""

from __future__ import annotations

import argparse
import re
import shutil
from collections import Counter
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent
EXCEL = BASE / "TeaWithNessie_GuestData_NiyatThapa.xlsx"
TEMPLATE = BASE / "TeaWithNessie_Invitation_Template_NiyatThapa.docx"
OUTPUT = BASE / "output"
FIELD_RE = re.compile(r"<<([A-Za-z0-9_]+)>>")

REQUIRED = {
    "guest_id", "first_name", "last_name", "courtesy_title", "email",
    "street_address", "town", "postcode", "venue_name", "venue_town",
    "venue_promise", "tea_date", "tea_time", "meeting_point",
    "rsvp_deadline", "cookie_flavour", "boat_tour_opt_in", "rsvp_email",
}


def text(value: object) -> str:
    return "" if value is None else str(value).strip()


def load_guests() -> tuple[list[str], list[dict[str, str]]]:
    """Read the Guests sheet and validate its headings and records."""
    if not EXCEL.exists():
        raise FileNotFoundError(f"Missing workbook: {EXCEL.name}")

    book = load_workbook(EXCEL, read_only=True, data_only=True)
    try:
        if "Guests" not in book.sheetnames:
            raise ValueError("The workbook needs a Guests sheet.")

        rows = book["Guests"].iter_rows(values_only=True)
        headers = [text(value) for value in next(rows, ())]
        if not headers or any(not header for header in headers):
            raise ValueError("The Guests header row is empty or contains blanks.")
        if len(headers) != len(set(headers)):
            raise ValueError("The Guests sheet contains duplicate headings.")

        missing = sorted(REQUIRED - set(headers))
        if missing:
            raise ValueError("Missing required columns: " + ", ".join(missing))

        guests, ids = [], set()
        for row_number, row in enumerate(rows, start=2):
            values = [text(value) for value in row]
            if not any(values):
                continue
            guest = dict(zip(headers, values))
            blank = sorted(field for field in REQUIRED if not guest.get(field))
            if blank:
                raise ValueError(
                    f"Guests row {row_number} has blank required values: "
                    + ", ".join(blank)
                )
            if guest["guest_id"] in ids:
                raise ValueError(f"Duplicate guest_id: {guest['guest_id']}")
            ids.add(guest["guest_id"])
            guests.append(guest)
    finally:
        book.close()

    if not guests:
        raise ValueError("No guest records were found.")
    return headers, guests


def table_paragraphs(table):
    for row in table.rows:
        for cell in row.cells:
            yield from cell.paragraphs
            for nested in cell.tables:
                yield from table_paragraphs(nested)


def story_paragraphs(story):
    yield from story.paragraphs
    for table in story.tables:
        yield from table_paragraphs(table)


def all_paragraphs(doc):
    yield from story_paragraphs(doc)
    for section in doc.sections:
        yield from story_paragraphs(section.header)
        yield from story_paragraphs(section.footer)


def template_fields() -> set[str]:
    if not TEMPLATE.exists():
        raise FileNotFoundError(f"Missing template: {TEMPLATE.name}")
    doc = Document(TEMPLATE)
    return {
        field
        for paragraph in all_paragraphs(doc)
        for field in FIELD_RE.findall(paragraph.text)
    }


def validate_template(headers: list[str]) -> None:
    fields = template_fields()
    if not fields:
        raise ValueError("The Word template contains no <<field_name>> placeholders.")
    missing = sorted(fields - set(headers))
    if missing:
        raise ValueError("Template fields missing from Excel: " + ", ".join(missing))


def replace_fields(paragraph, guest: dict[str, str]) -> None:
    if "<<" not in paragraph.text:
        return
    merged = FIELD_RE.sub(lambda m: guest.get(m.group(1), f"<<MISSING:{m.group(1)}>>"), paragraph.text)
    if paragraph.runs:
        paragraph.runs[0].text = merged
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(merged)


def merge_letter(guest: dict[str, str], destination: Path) -> None:
    doc = Document(TEMPLATE)
    for paragraph in all_paragraphs(doc):
        replace_fields(paragraph, guest)
    leftovers = [p.text for p in all_paragraphs(doc) if "<<" in p.text]
    if leftovers:
        raise ValueError(f"Unresolved fields for {guest['guest_id']}: {leftovers}")
    doc.save(destination)


def safe(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", value).strip("_") or "unknown"


def choose_guests(guests: list[dict[str, str]], venue: str | None):
    if not venue:
        return guests
    target = venue.casefold()
    selected = [
        guest for guest in guests
        if target in guest["venue_name"].casefold()
        or target in guest["venue_town"].casefold()
    ]
    if not selected:
        raise ValueError(f"No guests matched venue filter: {venue}")
    return selected


def write_summary(guests: list[dict[str, str]], files: list[str]) -> None:
    counts = Counter(guest["venue_name"] for guest in guests)
    lines = [
        "Tea with Nessie merge summary",
        "==============================",
        f"Letters created: {len(guests)}",
        "",
        "Venue counts:",
        *[f"- {venue}: {count}" for venue, count in sorted(counts.items())],
        "",
        "Generated files:",
        *[f"- {name}" for name in files],
    ]
    (OUTPUT / "merge_summary.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Create Tea with Nessie invitations.")
    parser.add_argument("--venue", help="Optional venue name or town filter, such as Dores.")
    args = parser.parse_args()

    headers, guests = load_guests()
    validate_template(headers)
    selected = choose_guests(guests, args.venue)

    if OUTPUT.exists():
        shutil.rmtree(OUTPUT)
    OUTPUT.mkdir()

    files = []
    for guest in selected:
        name = (
            f"invitation_{safe(guest['guest_id'])}_"
            f"{safe(guest['first_name'])}_{safe(guest['last_name'])}.docx"
        )
        merge_letter(guest, OUTPUT / name)
        files.append(name)
        print(f"Created: {name}")

    write_summary(selected, files)
    counts = Counter(guest["venue_name"] for guest in selected)
    print(f"\nMerge complete: {len(selected)} letters")
    for venue, count in sorted(counts.items()):
        print(f"  {venue}: {count}")
    print(f"Output folder: {OUTPUT}")


if __name__ == "__main__":
    main()
