import argparse
from spreadsheet import Spreadsheet


def parse_arguments():
    description = """
    This is a simple spreadsheet application that allows you to create, edit, and save spreadsheets.
    You can perform various operations such as entering data, applying formulas, formatting cells, 
    and saving your work in different file formats like JSON, YAML, Excel, CSV, and PDF.
    """
    parser = argparse.ArgumentParser(description=description)
    return parser.parse_args()


def main():
    parse_arguments()

    # Start a new spreadsheet application
    spreadsheet = Spreadsheet()
    spreadsheet.start_spreadsheet()


if __name__ == '__main__':
    main()


