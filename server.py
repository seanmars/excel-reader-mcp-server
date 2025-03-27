#!/usr/bin/env python3
import os
import json
from pathlib import Path
import pandas as pd
from mcp.server.fastmcp import FastMCP

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
def read_excel(filename: str, sheet_name: str = None) -> str:
    """Read an Excel file and return its contents as JSON

    Args:
        file_path: Path to the Excel file
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

        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        # Convert to JSON string, handling NaN/NaT values
        return df.to_json(orient='records', date_format='iso')
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def add(a: int, b: int) -> int:
    """Add two numbers"""
    return a + b

# Add a dynamic greeting resource


@mcp.resource("greeting://{name}")
def get_greeting(name: str) -> str:
    """Get a personalized greeting"""
    return f"Hello, {name}!"


if __name__ == "__main__":
    from dotenv import load_dotenv
    load_dotenv()

    res = get_res_folders()
    print(res)
    response = read_excel("PC_2025-03-27.xls", "2025-03-27")
    print(response)
