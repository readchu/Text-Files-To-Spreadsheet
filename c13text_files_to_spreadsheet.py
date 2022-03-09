#! python3
# c13text_files_to_spreadsheet.py - read in the contents of
# several text files (for now, in CWD) and insert those contents
# into a spreadsheet, with one line of text per row.

import sys, openpyxl
from pathlib import Path
import logging

logging.basicConfig(
    filename="c13TFTS_DEBUG.txt",
    level=logging.DEBUG,
    format=" %(asctime)s - %(levelname)s - %(message)s",
)
logging.disable(logging.CRITICAL)

def lines_to_sheet():
    wb = openpyxl.Workbook()
    sheet = wb.active

    f = Path.cwd()

    file_list = list(f.glob("*.txt"))

    for file in file_list:
        file_position = file_list.index(file) + 1
        sheet.cell(row=1, column=file_position).value = file.name
        open_file = open(file)
        file_lines = open_file.readlines()
        for line in file_lines:
            sheet.cell(
                row=file_lines.index(line) + 2, column=file_position
            ).value = line
            #alternative value=text[line] instead of ().value = line
    filename = "Separated_lines_from_" + Path.cwd().name + ".xlsx"
    try:
        wb.save(filename)
    except PermissionError:
        print(f"Saving {filename} not permitted. Terminated...")
        sys.exit()
