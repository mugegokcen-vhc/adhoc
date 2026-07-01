from pathlib import Path
import calendar
import csv
import re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


PAGEVIEW_HEADERS = ["Month", "url", "total_pageview"]
SCROLL_HEADERS = ["Month", "url", "scroll depth", "number of visitors", "drop off pct"]


def read_csv_rows(file_path: Path) -> list[list[str]]:
    """Read Clarity CSV with BOM-safe UTF-8 handling."""
    with file_path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.reader(f))


def get_metadata_value(rows: list[list[str]], key: str) -> str | None:
    """Find metadata values such as Date range, Page views, Metric, etc."""
    key_norm = key.strip().lower()
    for row in rows:
        if len(row) >= 2 and row[0].strip().lower() == key_norm:
            return row[1].strip()
    return None


def parse_date_range(date_range_text: str) -> str:
    """
    Convert Clarity date range into the month name expected in the output.

    Example input:
    06/01/2026 12:00 AM - 06/29/2026 11:59 PM

    Output:
    june
    """
    if not date_range_text:
        raise ValueError("Date range is missing")

    start_text = date_range_text.split(" - ", 1)[0].strip()
    start_dt = datetime.strptime(start_text, "%m/%d/%Y %I:%M %p")
    return calendar.month_name[start_dt.month].lower()


def regex_to_url(url_regex: str) -> str:
    r"""
    Convert the Clarity 'Visited URL matches regex' value into the real URL.

    Example input:
    ^https://info-prodotto\.it/integratori-di-magnesio/(\?.*)?$

    Output:
    https://info-prodotto.it/integratori-di-magnesio/
    """
    if not url_regex:
        raise ValueError("Visited URL matches regex is missing")

    url = url_regex.strip()

    if url.startswith("^"):
        url = url[1:]
    if url.endswith("$"):
        url = url[:-1]

    # Remove optional query-string regex suffix such as (\?.*)?
    url = re.sub(r"\(\\\?\.\*\)\??$", "", url)

    # Unescape common URL characters from regex notation.
    url = url.replace(r"\.", ".")
    url = url.replace(r"\/", "/")
    url = url.replace(r"\-", "-")
    url = url.replace(r"\_", "_")

    return url


def clean_int(value: str) -> int:
    return int(str(value).replace(",", "").strip())


def clean_float(value: str) -> float:
    return float(str(value).replace(",", ".").replace("%", "").strip())


def find_scroll_header_index(rows: list[list[str]]) -> int:
    """Find the row index where the scroll table starts."""
    for index, row in enumerate(rows):
        normalized = [cell.strip().lower() for cell in row]
        if (
            len(normalized) >= 3
            and normalized[0] == "scroll depth"
            and normalized[1] == "no. of visitors"
            and normalized[2] == "% drop off"
        ):
            return index

    raise ValueError("Scroll table header could not be found")


def parse_clarity_csv(file_path: Path) -> tuple[list, list[list]]:
    """Parse one Clarity Scroll CSV into one pageview row and multiple scroll rows."""
    rows = read_csv_rows(file_path)

    date_range = get_metadata_value(rows, "Date range")
    url_regex = get_metadata_value(rows, "Visited URL matches regex")
    page_views = get_metadata_value(rows, "Page views")
    metric = get_metadata_value(rows, "Metric")

    if metric and metric.strip().lower() != "scroll":
        raise ValueError(f"Metric is not Scroll. Found: {metric}")

    month = parse_date_range(date_range)
    url = regex_to_url(url_regex)
    total_pageview = clean_int(page_views)

    pageview_row = [month, url, total_pageview]

    header_index = find_scroll_header_index(rows)
    scroll_rows = []

    for row in rows[header_index + 1:]:
        if len(row) < 3 or not row[0].strip():
            continue

        scroll_rows.append([
            month,
            url,
            clean_int(row[0]),
            clean_int(row[1]),
            clean_float(row[2]),
        ])

    return pageview_row, scroll_rows


def deduplicate(rows: list[list]) -> list[list]:
    """Remove duplicate rows while preserving order."""
    seen = set()
    unique_rows = []

    for row in rows:
        key = tuple(row)
        if key not in seen:
            seen.add(key)
            unique_rows.append(row)

    return unique_rows


def autofit_basic(ws):
    """Apply practical column widths."""
    for column_cells in ws.columns:
        column_letter = get_column_letter(column_cells[0].column)
        max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
        width = min(max(max_length + 2, 12), 70)
        ws.column_dimensions[column_letter].width = width


def style_sheet(ws, table_name: str):
    """Basic readable Excel formatting."""
    header_fill = PatternFill("solid", fgColor="0F766E")
    header_font = Font(bold=True, color="FFFFFF")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    autofit_basic(ws)

    max_row = ws.max_row
    max_col = ws.max_column

    if max_row > 1:
        ref = f"A1:{get_column_letter(max_col)}{max_row}"
        table = Table(displayName=table_name, ref=ref)
        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        table.tableStyleInfo = style
        ws.add_table(table)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top")


def build_excel(input_dir: Path, output_file: Path):
    pageview_rows = []
    scroll_rows = []
    errors = []

    csv_files = sorted(input_dir.glob("*.csv"))

    if not csv_files:
        raise FileNotFoundError(f"No CSV files found in: {input_dir}")

    for csv_file in csv_files:
        try:
            pageview_row, file_scroll_rows = parse_clarity_csv(csv_file)
            pageview_rows.append(pageview_row)
            scroll_rows.extend(file_scroll_rows)
            print(f"Processed: {csv_file.name}")
        except Exception as exc:
            errors.append((csv_file.name, str(exc)))
            print(f"Skipped: {csv_file.name} | Reason: {exc}")

    pageview_rows = deduplicate(pageview_rows)
    scroll_rows = deduplicate(scroll_rows)

    wb = Workbook()
    pageview_ws = wb.active
    pageview_ws.title = "Pageviews"
    scroll_ws = wb.create_sheet("Scroll")

    pageview_ws.append(PAGEVIEW_HEADERS)
    for row in pageview_rows:
        pageview_ws.append(row)

    scroll_ws.append(SCROLL_HEADERS)
    for row in scroll_rows:
        scroll_ws.append(row)

    style_sheet(pageview_ws, "PageviewsTable")
    style_sheet(scroll_ws, "ScrollTable")

    output_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_file)

    print("\nDone")
    print(f"Output file: {output_file}")
    print(f"Pageview rows: {len(pageview_rows)}")
    print(f"Scroll rows: {len(scroll_rows)}")

    if errors:
        print("\nFiles skipped:")
        for file_name, reason in errors:
            print(f"- {file_name}: {reason}")


def main():
    project_root = Path(__file__).resolve().parents[1]

    input_dir = project_root / "data"
    output_file = project_root / "output" / "clarity_combined_output.xlsx"

    build_excel(input_dir, output_file)


if __name__ == "__main__":
    main()
