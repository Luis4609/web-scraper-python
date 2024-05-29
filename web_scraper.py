"""Module with the web scraping functionalities"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from typing import List, Optional, Tuple, Any


def fetch_webpage(url: str) -> Optional[str]:
    """
    Fetch the HTML content of a webpage.

    Args:
        url (str): The URL of the webpage to fetch.

    Returns:
        Optional[str]: The HTML content of the webpage, or None if the request failed.
    """
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Raise an exception for HTTP errors
        return response.text
    except requests.RequestException as e:
        print(f"Failed to fetch the web page: {e}")
        return None

def parse_html(html_content: str) -> BeautifulSoup:
    """
    Parse HTML content using BeautifulSoup.

    Args:
        html_content (str): The HTML content to parse.

    Returns:
        BeautifulSoup: The parsed BeautifulSoup object.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    return soup

def extract_table_data(table: BeautifulSoup) -> List[List[str]]:
    """
    Extract data from an HTML table.

    Args:
        table (BeautifulSoup): The BeautifulSoup object representing the table.

    Returns:
        List[List[str]]: A list of rows, where each row is a list of cell values.
    """
    rows = table.find_all('tr')
    table_data = []

    for row in rows:
        cols = row.find_all(['td', 'th'])
        cols = [col.text.strip() for col in cols]
        table_data.append(cols)

    return table_data

def create_excel_workbook() -> Tuple[Workbook, Worksheet]:
    """
    Create an Excel workbook and worksheet.

    Returns:
        Tuple[Workbook, Worksheet]: The created workbook and worksheet.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Table Data"
    return wb, ws

def add_data_to_sheet(ws: Worksheet, df: pd.DataFrame) -> None:
    """
    Add DataFrame data to an Excel worksheet.

    Args:
        ws (Worksheet): The Excel worksheet.
        df (pd.DataFrame): The DataFrame containing the data to add.
    """
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            apply_cell_formatting(cell, r_idx)

def apply_cell_formatting(cell: Any, row_idx: int) -> None:
    """
    Apply formatting to a cell.

    Args:
        cell (Any): The cell to format.
        row_idx (int): The row index of the cell.
    """
    if row_idx == 1:  # Apply formatting to header
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    else:
        cell.alignment = Alignment(horizontal="left")

def adjust_column_widths(ws: Worksheet) -> None:
    """
    Adjust the widths of columns in an Excel worksheet based on content.

    Args:
        ws (Worksheet): The Excel worksheet.
    """
    for col in ws.columns:
        max_length = max((len(str(cell.value)) for cell in col if cell.value), default=0)
        adjusted_width = max_length + 2
        column_letter = col[0].column_letter
        ws.column_dimensions[column_letter].width = adjusted_width

def export_to_excel(data: List[List[str]], filename: str) -> None:
    """
    Export table data to an Excel file with formatting.

    Args:
        data (List[List[str]]): The table data to export.
        filename (str): The name of the Excel file to create.
    """
    df = pd.DataFrame(data[1:], columns=data[0])  # Assuming the first row is the header

    wb, ws = create_excel_workbook()
    add_data_to_sheet(ws, df)
    adjust_column_widths(ws)

    wb.save(filename)

def main() -> None:
    """
    Main function to fetch a webpage, extract table data, and export it to an Excel file.
    """
    url = 'https://www.ine.es/dyngs/INEbase/es/operacion.htm?c=Estadistica_C&cid=1254736176802&menu=ultiDatos&idp=1254735976607'  # Replace with the url you want to scrap
    table_class = 'tablaCat'  # Replace with the actual class of the table

    html_content = fetch_webpage(url)
    if html_content:
        soup = parse_html(html_content)
        table = soup.find('table', class_=table_class)

        if table:
            print("Table found!")
            data = extract_table_data(table)

            # Export the DataFrame to an Excel file with formatting
            export_to_excel(data, 'table_data.xlsx')
            print("Data has been exported to table_data.xlsx")
        else:
            print("Table not found.")

if __name__ == "__main__":
    main()
