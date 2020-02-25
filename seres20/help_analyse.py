import xlrd
from xlutils.copy import copy
import datetime
import os
import re
import sys
import platform
from typing import List, Any

USAGE = 'Usage: python3 exclude_further.py <input xls path> <username>'

# Configuration
PAPER_MIN_LENGTH = 4
TITLE_ROW_HEADER = ''
ABSTRACT_ROW_HEADER = ''
AN_ROW_HEADER = ''
PAGE_ROW_HEADER = ''


def extract_pages(s: str) -> int:
    """Extracts page amount from column string data"""

    pages = s.strip()
    single_digit = re.search(r'^(\d+)$', pages)
    double_digit = re.search(r'(\d+)-(\d+)', pages)

    if single_digit:
        return 1
    elif double_digit:
        multi_lines = pages.split(',')
        if len(multi_lines) > 1:
            return sum([extract_pages(x) for x in multi_lines])
        else:
            return len(range(int(double_digit.group(1)), int(double_digit.group(2)) + 1))
    else:
        return 0


def read_previously_excluded() -> List[int]:
    """Reads previously excluded rows from file"""
    with open('saved_exclude.txt', 'r') as f:
        return [x.strip() for x in f.readlines() if len(x) > 0]


def read_previously_included() -> List[int]:
    """Reads previously included rows from file"""
    with open('saved_include.txt', 'r') as f:
        return [x.strip() for x in f.readlines() if len(x) > 0]


def save_work(included_ans: List[str], excluded_ans: List[str]) -> None:
    """Save work to text files for later use"""
    with open('saved_exclude.txt', 'w') as f:
        for i in excluded_ans:
            f.write(str(i) + '\n')

    with open('saved_include.txt', 'w') as f:
        for i in included_ans:
            f.write(str(i) + '\n')

    print('\nSave completed!\n')


def get_an(ws: Any, col_index: dict, row: int) -> str:
    """Get accession number of row item"""
    return str(ws.cell_value(row, col_index[AN_ROW_HEADER]))


def get_row(ws: Any, n_rows: int, col_index: dict, an: str) -> int:
    """Get row number of accession number item"""
    for i in range(n_rows):
        if str(ws.cell_value(i, col_index[AN_ROW_HEADER])) == an:
            return i
    return -1


if __name__ == "__main__":
    # Usage
    if len(sys.argv) != 3:
        print(len(sys.argv))
        print(USAGE)
        sys.exit(1)

    # Configuration check
    if (len(TITLE_ROW_HEADER) == 0 or len(ABSTRACT_ROW_HEADER) == 0 or
            len(AN_ROW_HEADER) == 0 or len(PAGE_ROW_HEADER) == 0):
        print('Configure proper title headers to source code and try again')
        sys.exit(1)

    file_path = sys.argv[1]
    user = sys.argv[2]
    rationale = f'{user} (rationale)'

    read_wb = xlrd.open_workbook(file_path)
    read_ws = read_wb.sheet_by_index(0)

    n_rows = read_ws.nrows
    col_index = {}

    excluded_ans = []
    included_ans = []
    too_short = []

    # Prepare row titles
    for num, cell in enumerate(read_ws.row(0)):
        col_index[cell.value] = num

    # Add user row title and rationale title
    col_index[user] = max(col_index.values()) + 1
    col_index[rationale] = col_index[user] + 1

    # Platform check
    if platform.system() == 'Linux':
        def clear(): return os.system('clear')
    elif platform.system() == 'Windows':
        def clear(): return os.system('cls')
    else:
        print('Apple not currently supported')
        sys.exit(1)

    # Load previous saves
    if(os.path.isfile('saved_exclude.txt')):
        excluded_ans.extend(read_previously_excluded())

    if(os.path.isfile('saved_include.txt')):
        included_ans.extend(read_previously_included())

    # Remove duplicates based on title
    ans_encountered = []
    titles = []
    excluded_title_rows = []

    # Process titles in loaded files
    for an in excluded_ans:
        titles.append(read_ws.cell_value(
            get_row(read_ws, n_rows, col_index, an), col_index[TITLE_ROW_HEADER]))

    for an in included_ans:
        titles.append(read_ws.cell_value(
            get_row(read_ws, n_rows, col_index, an), col_index[TITLE_ROW_HEADER]))

    # Evaluate titles
    for row, i in enumerate(read_ws.col_values(col_index[TITLE_ROW_HEADER])[1:], start=1):
        if i not in titles:
            titles.append(i)
        elif get_an(read_ws, col_index, row) not in included_ans and \
                get_an(read_ws, col_index, row) not in excluded_ans:
            excluded_title_rows.append(row)

    # Exclusion based on paper length
    for row, i in enumerate(read_ws.col_values(col_index[PAGE_ROW_HEADER])[1:], start=1):
        if get_an(read_ws, col_index, row) not in excluded_ans and get_an(read_ws, col_index, row) not in included_ans:
            if extract_pages(i) < PAPER_MIN_LENGTH:
                too_short.append(get_an(read_ws, col_index, row))

    # Abstract exclusion/inclusion functionality
    for row, i in enumerate(read_ws.col_values(col_index[ABSTRACT_ROW_HEADER])[1:], start=1):
        if (get_an(read_ws, col_index, row) not in excluded_ans and
                get_an(read_ws, col_index, row) not in included_ans and
                row not in excluded_title_rows and
                get_an(read_ws, col_index, row) not in too_short):

            clear()

            print('Welcome to evaluate work by abstract feature!')
            print(
                f'\n{n_rows - 1 - len(included_ans) - len(excluded_ans) - len(excluded_title_rows) - len(too_short)}' +
                ' more abstracts to go!\n\n'
            )
            print(f'{i.strip()}\n')

            include = ''

            while include != 'yes' and include != 'no':
                if include == 'print':
                    print(
                        f'Currently excluded access_numbers by abstract: {excluded_ans}')
                    print(
                        f'Currently included access_numbers by abstract: {included_ans}\n')

                elif include == 'save':
                    save_work(included_ans, excluded_ans)
                elif include == 'quit':
                    print('Bye!')
                    sys.exit(0)

                include = input('Include the work (yes/no/print/save/quit): ')

            if include == 'no':
                excluded_ans.append(get_an(read_ws, col_index, row))
            if include == 'yes':
                included_ans.append(get_an(read_ws, col_index, row))

    # Automatical save on abstract evaluation exit
    save_work(included_ans, excluded_ans)

    # Saving info to .xls
    wb_copy = copy(read_wb)
    write_sheet = wb_copy.get_sheet(0)

    output_file = f'{file_path[:-4]}_edited_{datetime.datetime.today().strftime("%Y%m%d")}.xls'  # noqa: E228,E999

    # Add row headers
    write_sheet.write(0, col_index[user], user)
    write_sheet.write(
        0, col_index[rationale], rationale)

    # Write decision
    for row, i in enumerate(read_ws.col_values(0)[1:], start=1):
        if (get_an(read_ws, col_index, row) in included_ans
                and row not in excluded_title_rows
                and row not in too_short):
            write_sheet.write(row, col_index[user], 'YES')
        else:
            write_sheet.write(row, col_index[user], 'NO')

    # Write rationale
    for row, i in enumerate(read_ws.col_values(0)[1:], start=1):
        if get_an(read_ws, col_index, row) in too_short:
            write_sheet.write(
                row, col_index[rationale], 'Paper is too short (under 4 pages)')
        elif row in excluded_title_rows:
            write_sheet.write(
                row, col_index[rationale], 'Title was duplicate')
        else:
            write_sheet.write(
                row, col_index[rationale], 'Decision based on abstract')

    wb_copy.save(output_file)
    print(f'\nWork has been saved to file {output_file}!')
