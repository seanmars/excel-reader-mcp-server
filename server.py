#!/usr/bin/env python3
import os
import json
from pathlib import Path
from typing import Any

import pandas as pd
from mcp.server.fastmcp import FastMCP
from pandas import DataFrame
from pydantic.v1.utils import to_lower_camel

# Create an MCP server
mcp = FastMCP("excel-reader", dependencies=["pandas", "openpyxl", "xlrd"])


@mcp.tool()
def get_res_folders() -> list[Path]:
    """
    Get the environment variable for resource folders
    env key = MCP_RESOURCE_FOLDERS
    Returns:
        List of resource folders
    """
    res_folders = os.getenv("MCP_RESOURCE_FOLDERS")
    if res_folders:
        return [Path(folder.strip()).resolve() for folder in res_folders.split(",")]

    return []


@mcp.tool()
def get_excel_file_path(filename: str) -> str:
    """
    Get the path to an Excel file in the resource folders
    Args:
        filename: Name of the Excel file

    Returns:
        Path to the Excel file
    """
    res_folders = get_res_folders()
    for folder in res_folders:
        # Check if the folder is ending with a slash remove it
        if str(folder).endswith("/"):
            folder = folder[:-1]
        # Check if the folder exists
        if not folder.exists():
            continue
        file_path = f"{folder}/{filename}"
        if Path(file_path).exists():
            return file_path
    return ""


@mcp.tool()
def fetch_sheet_names(filename: str) -> str | list[Any]:
    try:
        file_path = get_excel_file_path(filename)
        if not file_path:
            raise FileNotFoundError(
                f"File {filename} not found in resource folders.")

        return pd.ExcelFile(file_path).sheet_names
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def read_game_data(filename) -> str:
    """Read game data from an Excel file and return it as JSON
    """
    try:
        file_path = get_excel_file_path(filename)
        print(f"File path: {file_path}")
        if not file_path:
            raise FileNotFoundError(
                f"File {filename} not found in resource folders.")

        sheets = fetch_sheet_names(filename)
        sheet_name = sheets[0] if isinstance(sheets, list) else sheets
        print(f"Sheet name: {sheet_name}")
        # Check if the sheet name is valid
        if not sheet_name:
            raise ValueError("Sheet name is empty.")

        # Read the Excel file
        xl = pd.ExcelFile(file_path)

        # parse first 20 row to get row number which value == 'type' in first column
        df: DataFrame = xl.parse(sheet_name, header=None, nrows=20)
        row = df.iloc[:, 0].tolist().index('type')
        header_row = row - 1
        type_row = row
        info_row = row + 1
        skip_rows = [type_row, info_row]
        end_col = df.iloc[row, :].tolist().index('###')
        print(f"header_row: {header_row}, type_row: {type_row}, info_row: {info_row}, end_col: {end_col}")
        df = xl.parse(sheet_name, header=header_row,
                      skiprows=lambda x: x in skip_rows,
                      usecols=list(range(0, end_col)))
        # remove row if first column is 'ps' case-insensitive
        df = df[df.iloc[:, 0].str.lower() != 'ps']

        # find first row which value == '###' in first column
        end_row = df.iloc[:, 0].tolist().index('###')
        print(f"end_row: {end_row}")
        df = df.iloc[:end_row - 1, :]

        # print total rows
        print(f"Total rows: {len(df.index)}")

        # Convert to JSON string, handling NaN/NaT values
        with open('output.json', 'w', encoding='utf-8') as file:
            df.to_json(file, orient='records', date_format='iso', force_ascii=False)
        return df.to_json(orient='records', date_format='iso', force_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def read_excel(filename: str, sheet_name: str | None = None) -> str:
    """Read an Excel file and return its contents as JSON

    Args:
        filename: Path to the Excel file
        sheet_name: Name of the sheet to read (optional, defaults to first sheet)

    Returns:
        JSON string containing the Excel data
    """
    try:
        file_path = get_excel_file_path(filename)
        print(f"File path: {file_path}")
        if not file_path:
            raise FileNotFoundError(
                f"File {filename} not found in resource folders.")

        sheets = fetch_sheet_names(filename)
        if not sheet_name:
            sheet_name = sheets[0] if isinstance(sheets, list) else sheets
        # Check if the sheet name is valid
        if not sheet_name:
            raise ValueError("Sheet name is empty.")

        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        # Convert to JSON string, handling NaN/NaT values
        return df.to_json(orient='records', date_format='iso')
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def get_excel_file_list():
    res_folders = get_res_folders()
    file_list = []
    for folder in res_folders:
        print(folder)
        for file in folder.glob("*.xlsx"):
            file_list.append(str(file))
        for file in folder.glob("*.xls"):
            file_list.append(str(file))
    return file_list


if __name__ == "__main__":
    from dotenv import load_dotenv

    load_dotenv()

    filename = 'ITEM.xlsx'
    result = read_game_data(filename)
    # print(result)
    print('done')
