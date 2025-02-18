# Excel MCP Server

[![smithery badge](https://smithery.ai/badge/@negokaz/excel-mcp-server)](https://smithery.ai/server/@negokaz/excel-mcp-server)
[![NPM Version](https://img.shields.io/npm/v/excel-mcp-server)](https://www.npmjs.com/package/excel-mcp-server)

A Model Context Protocol (MCP) server that reads and writes spreadsheet data to MS Excel file.

## Features

- Read text values from MS Excel file
- Write text values to MS Excel file

## Requirements

- Node.js 20.x or later

## Supported file formats

- xlsx

## Installation

excel-mcp-server is automatically installed by adding the following configuration to the MCP servers configuration.

### Installing via Smithery

To install Excel MCP Server for Claude Desktop automatically via [Smithery](https://smithery.ai/server/@negokaz/excel-mcp-server):

```bash
npx -y @smithery/cli install @negokaz/excel-mcp-server --client claude
```

For Windows:
```json
{
  "mcpServers": {
    "excel": {
        "command": "cmd",
        "args": ["/c", "npx", "--yes", "excel-mcp-server"],
    }
}
```

For other platforms:
```json
{
  "mcpServers": {
    "excel": {
        "command": "npx",
        "args": ["--yes", "excel-mcp-server"],
    }
}
```

## License

Copyright (c) 2025 Kazuki Negoro

excel-mcp-server is released under the [MIT License](LICENSE)
