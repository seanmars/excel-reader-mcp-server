# Excel Reader MCP Server

## Description

This is a simple MCP server that reads Excel files and returns the data in a structured format. It uses the `openpyxl` library to read Excel files.

## Setup

- Create a `.env` file in the root directory of the project. You can use the `.env.example` file as a template.

## Install

### Install to Claude Desktop

```shell
uv run mcp install server.py -f .env
```