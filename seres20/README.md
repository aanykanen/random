# Seres20

During software engineering research course there was assignment to do a systematic literature review. To help the process of evaluating search results from .xls file a script was developed.

## Requirements
- Python 3.7+
- Packages:
    - xlrd
    - xlwt
    - xlutils

## Usage
1. Configure following parameters to script
    - PAPER_MIN_LENGTH
        - Minimum length of article
    - TITLE_ROW_HEADER
        - Cell value in row 0 and column where title information is located
    - ABSTRACT_ROW_HEADER
        - Cell value in row 0 and column where abstract information is located
    - AN_ROW_HEADER
        - Cell value in row 0 and column where access number information is located
    - PAGE_ROW_HEADER
        - Cell value in row 0 and column where pages information is located
2. Run script using following syntax
    - ```python3 help_analyse.py <path to .xls> <username>```

## Notes
- Username is only used to create named columns, this allows multiple users to use same script for same .xls sheet in case there are multiple authors
- Remember that the saved_include.txt and saved_exclude.txt files are always overwritten! If you have a set you don't wish to overwrite, copy the files elsewhere for later use.