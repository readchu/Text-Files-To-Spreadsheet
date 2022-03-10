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

#could write it instead to somehow ask user to specify a folder
#and also to just specify certain files instead

def lines_to_sheet():
    wb = openpyxl.Workbook()
    sheet = wb.active

    f = Path.cwd()

    file_list = list(f.glob('*.txt'))

    for fl_index, file in enumerate(file_list):
        file_position = fl_index + 1
        sheet.cell(row=1, column=file_position).value = file.name
        open_file = open(file)
        file_lines = open_file.readlines()
        for line_index, line in enumerate(file_lines):
            sheet.cell(
                row=line_index + 2, column=file_position
            ).value = line.strip()
            logging.debug(f'This is the active line: {line}')
            logging.debug(f'This is where it\'ll be placed: {sheet.cell(row=file_lines.index(line) + 2, column=file_position)}')
            #alternative value=text[line] instead of ().value = line
            #use .strip() to reduce whitespace in sheet

#    debug_column = []
#    for row in sheet['C']:
#        debug_column.append(row.value)
#    logging.debug(f'This is the contents of column C: {debug_column}')

    filename = f'Separated_lines_from_{f.name}.xlsx'
    try:
        wb.save(filename)
    except PermissionError:
        print(f'Saving {filename} not permitted. Terminated...')
        sys.exit()
