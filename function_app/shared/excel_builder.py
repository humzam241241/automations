from io import BytesIO
from typing import List

from openpyxl import Workbook


def build_workbook(headers: List[str], row: List[str]) -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(headers)
    worksheet.append(row)
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()
