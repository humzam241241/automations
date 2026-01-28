import csv
from io import StringIO
from typing import List


def generate_csv(headers: List[str], row: List[str], summary: str) -> bytes:
    buffer = StringIO()
    writer = csv.writer(buffer)
    writer.writerow(headers)
    writer.writerow(row)

    summary_row = [""] * len(headers)
    if headers:
        summary_row[0] = "SUMMARY"
    if "Summary" in headers:
        summary_row[headers.index("Summary")] = summary
    writer.writerow(summary_row)
    return buffer.getvalue().encode("utf-8")
