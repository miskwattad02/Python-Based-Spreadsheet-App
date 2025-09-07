import math
import re
import json
import numpy as np
import yaml
from openpyxl import Workbook
import csv
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
import pandas as pd
import pdfplumber


def number_to_excel_column(number):
    """
    Converts a number to its corresponding Excel column letter.

    :param number: The number to convert.
    :return: The Excel column letter corresponding to the number.
    """
    if number <= 0:
        return ''

    quotient, remainder = divmod(number - 1, 26)
    return number_to_excel_column(quotient) + chr(65 + remainder)


def _excel_column_to_number(column_str):
    """
    Converts an Excel column letter to its corresponding number.

    :param column_str: The Excel column letter.
    :return: The corresponding column number.
    """
    result = 0
    for i, char in enumerate(reversed(column_str)):
        result += (ord(char) - 64) * (26 ** i)
    return result


def _excel_to_indices(excel_str):
    """
    Converts an Excel-style cell reference to row and column indices.

    :param excel_str: The Excel-style cell reference (e.g., "A1").
    :return: A tuple containing the row index and column index.
    """
    column_str = ''.join(filter(str.isalpha, excel_str))
    row_str = ''.join(filter(str.isdigit, excel_str))

    column_index = _excel_column_to_number(column_str) - 1

    row_index = int(row_str) - 1

    return row_index, column_index


def _find_letter_number_indices(expression):
    """
    Finds letter-number pairs (cell references) in a given expression.

    :param expression: The expression to search.
    :return: A list of tuples containing the start and end indices of cell references in the expression.
    """
    pattern = re.compile(r'[a-zA-Z]+\d+')
    matches = pattern.finditer(expression)
    indices = [(match.start(), match.end() - 1) for match in matches]
    return indices


def _replace_coord_names(expression, indices, sheet_values):
    """
    Replaces cell coordinates in an expression with corresponding values from a sheet.

    :param expression: The expression containing cell coordinates.
    :param indices: A list of tuples containing the start and end indices of cell references.
    :param sheet_values: The values of cells in the sheet.
    :return: The expression with cell coordinates replaced by values.
    """
    new_expression = list(expression)
    offset = 0

    for first, last in indices:
        coord_name = expression[first:last + 1]
        i, j = _excel_to_indices(coord_name)
        replacement_value = sheet_values[i][j]
        new_expression[first + offset:last + 1 + offset] = list(replacement_value)
        offset += len(replacement_value) - (last - first + 1)

    new_expression = ''.join(new_expression)
    return new_expression


def solve_expression(expression, sheet_values):
    """
    Solves a mathematical expression with cell references.

    :param expression: The mathematical expression to solve.
    :param sheet_values: The values of cells in the sheet.
    :return: The result of the expression evaluation.
    """
    def average(*args):
        return sum(args) / len(args)

    def sum_(*args):
        return sum(args)

    def if_(*args):
        if args[0]:
            return args[1]
        return args[2]

    def sqrt(*args):
        return math.sqrt(sum(args))

    def countif_(*args):
        count = 0
        for i in range(len(args) - 1):
            if eval(str(args[i]) + args[-1]):
                count += 1
        return count

    expression = expression.upper().replace(" ", "")
    indices = _find_letter_number_indices(expression)
    expression = _replace_coord_names(expression, indices, sheet_values)
    sum_replace = expression.lower().replace("sum", "sum_")
    if_replace = sum_replace.replace("if", "if_")
    return eval(if_replace)

def _next_letter(input_letters):
    """
    Generates the next Excel column letter given a sequence of letters.

    :param input_letters: The input letters.
    :return: The next Excel column letter.
    """
    carry = 1
    result = []

    for letter in input_letters[::-1]:
        value = ord(letter) - ord('A') + carry
        carry = value // 26
        result.append(chr((value % 26) + ord('A')))

    if carry:
        result.append('A')

    return ''.join(result[::-1])


def _replace_coords_names(function, indices):
    """
    Replaces cell coordinates in a function with the next column letters.

    :param function: The function containing cell coordinates.
    :param indices: A list of tuples containing the start and end indices of cell references.
    :return: The function with cell coordinates replaced by the next column letters.
    """
    new_expression = list(function)
    offset = 0
    for first, last in indices:
        coord_name = function[first:last + 1]
        column_str = ''.join(filter(str.isalpha, coord_name))
        row_str = ''.join(filter(str.isdigit, coord_name))
        letter = _next_letter(column_str)
        new_coord_name = letter+row_str
        new_expression[first + offset:last + 1 + offset] = list(new_coord_name)
        offset += len(new_coord_name) - (last - first + 1)
    new_expression = ''.join(new_expression)
    return new_expression


def get_next_function(function):
    """
    Generates the next function with updated cell coordinates.

    :param function: The current function containing cell coordinates.
    :return: The next function with updated cell coordinates.
    """
    indices = _find_letter_number_indices(function)
    return _replace_coords_names(function, indices)


# ####################################### Save/Open Files ####################################### #

def read_json_file(file_name):
    """
    Reads data from a JSON file.

    :param file_name: The path to the JSON file.
    :return: The data read from the JSON file.
    """
    with open(file_name) as f:
        data = json.load(f)
    return data


def read_yaml_file(file_name):
    """
    Reads data from a YAML file.

    :param file_name: The path to the YAML file.
    :return: The data read from the YAML file.
    """
    with open(file_name) as f:
        data = yaml.safe_load(f)
    return data


def read_excel_file(file_name):
    """
    Reads data from a Excel file.

    :param file_name: The path to the Excel file.
    :return: The data read from the Excel file.
    """
    df = pd.read_excel(file_name, header=None)
    df = df.replace({np.nan: None})
    data = df.values.tolist()
    return data


def read_csv_file(file_name):
    """
    Reads data from a CSV file.

    :param file_name: The path to the CSV file.
    :return: The data read from the CSV file.
    """
    df = pd.read_csv(file_name, header=None)
    df = df.replace({np.nan: None})
    data = df.values.tolist()
    return data


def read_pdf_file(file_name):
    """
    Reads data from a PDF file.

    :param file_name: The path to the PDF file.
    :return: The data read from the PDF file.
    """
    with pdfplumber.open(file_name) as pdf:
        data = []
        for page in pdf.pages:
            page_text = page.extract_text()
            lines = page_text.split("\n")
            lines = [line.strip() for line in lines if line.strip()]
            rows = [[None if field == 'None' else field for field in line.split()] for line in lines]
            data.extend(rows)
    return data


def write_yaml_file(file_name, data):
    """
    Writes workbook data to a YAML file.

    :param file_name: The name of the YAML file to write.
    :param data: The workbook data to be written.
    :return: None
    """
    with open(file_name, 'w') as f:
        yaml.dump(data, f)


def write_json_file(file_name, data):
    """
    Writes workbook data to a JSON file.

    :param file_name: The name of the JSON file to write.
    :param data: The workbook data to be written.
    :return: None
    """
    with open(file_name, 'w') as f:
        json.dump(data, f)


def write_excel_file(file_name, data):
    """
    Writes workbook data to an Excel file.

    :param file_name: The name of the Excel file to write.
    :param data: The workbook data to be written.
    :return: None
    """
    wb = Workbook()
    ws = wb.active
    for row_index, row_data in enumerate(data, start=1):
        for col_index, cell_value in enumerate(row_data, start=1):
            ws.cell(row=row_index, column=col_index, value=cell_value)
    wb.save(file_name)


def write_csv_file(file_name, data):
    """
    Writes workbook data to a CSV file.

    :param file_name: The name of the CSV file to write.
    :param data: The workbook data to be written.
    :return: None
    """
    with open(file_name, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerows(data)


def write_pdf_file(file_name, data):
    """
    Writes workbook data to a PDF file.

    :param file_name: The name of the PDF file to write.
    :param data: The workbook data to be written.
    :return: None
    """
    c = canvas.Canvas(file_name, pagesize=landscape(letter))
    cell_width = 50
    cell_height = 20
    for i, row in enumerate(data):
        for j, cell in enumerate(row):
            x = 50 + j * cell_width
            y = 550 - (i * cell_height)
            c.drawString(x, y, str(cell))
    c.save()


