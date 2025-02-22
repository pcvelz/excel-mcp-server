#!/usr/bin/env node
import * as pkg from '../package.json';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
  CallToolResult,
} from '@modelcontextprotocol/sdk/types.js';
import fs from 'fs';
import ExcelJS from 'exceljs';
import { z, ZodError } from 'zod';
import { zodToJsonSchema } from 'zod-to-json-schema';

class ExcelMcpServer {

  private server: Server;

  private workbookCache = new Map<string, ExcelJS.Workbook>();

  private zFileAbsolutePath =
    z.string().describe('Absolute path to the Excel file');
  private zSheetName =
    z.string().describe('Sheet name in the Excel file');
  private zRange =
    z.string().describe('Range of cells in the Excel sheet (e.g., "A1:C10")');
  private zData =
    z.array(z.array(z.string())).describe('Data to write to the Excel sheet');

  private ReadSheetNameSchema = z.object({
    fileAbsolutePath: this.zFileAbsolutePath,
  });
  private ReadSheetDataSchema = z.object({
    fileAbsolutePath: this.zFileAbsolutePath,
    sheetName: this.zSheetName.optional(),
    range: this.zRange.optional(),
  });
  private WriteSheetDataSchema = z.object({
    fileAbsolutePath: this.zFileAbsolutePath,
    sheetName: this.zSheetName,
    range: this.zRange,
    data: this.zData,
  });

  constructor() {
    this.server = new Server({
      name: pkg.name,
      version: pkg.version,
      description: pkg.description,
    },
    {
      capabilities: {
        tools: {}
      }
    });
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'read_sheet_names',
          description: 'List all sheet names in an Excel file',
          inputSchema: zodToJsonSchema(this.ReadSheetNameSchema),
        },
        {
          name: 'read_sheet_data',
          description: 'Read data from the Excel sheet.' +
            'If there is a large number of data, it reads a part of data.' +
            'To read more data, adjust range parameter and make requests again.',
          inputSchema: zodToJsonSchema(this.ReadSheetDataSchema),
        },
        {
          name: 'write_sheet_data',
          description: 'Write data to the Excel sheet',
          inputSchema: zodToJsonSchema(this.WriteSheetDataSchema),
        },
      ]
    }));
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        switch (request.params.name) {
          case 'read_sheet_names':
            return this.handleReadSheetNames(this.ReadSheetNameSchema.parse(request.params.arguments));
          case 'read_sheet_data':
            return this.handleReadSheetData(this.ReadSheetDataSchema.parse(request.params.arguments));
          case 'write_sheet_data':
            return this.handleWriteSheetData(this.WriteSheetDataSchema.parse(request.params.arguments));
          default:
            throw new McpError(ErrorCode.MethodNotFound, `Tool [${request.params.name}] not found`);
        }
      } catch (error) {
        if (error instanceof McpError) {
          throw error;
        } else if (error instanceof ZodError) {
          throw new McpError(ErrorCode.InvalidParams, error.issues.map(e => `[${e.code}:${e.path}] ${e.message}`).join(', '));
        } else {
          throw new McpError(ErrorCode.InternalError, error instanceof Error ? error.message : 'Unknown error');
        }
      }
    });
  }

  public async start() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('MCP Server started');
  };

  /**
   * List all sheet names in an Excel file
   * @returns Sheet names
   */
  private async handleReadSheetNames(args: z.infer<typeof this.ReadSheetNameSchema>): Promise<CallToolResult> {
    const { fileAbsolutePath } = args;
    const workbook = await this.readWorkbook(fileAbsolutePath);
    const sheetNames = workbook.worksheets.map(sheet => sheet.name);
    return {
      content: sheetNames.map(name => ({
        type: 'text',
        text: name,
      }))
    }
  }

  /**
   * Read data from the Excel sheet
   * @returns Spreadsheet data in HTML table format
   */
  private async handleReadSheetData(args: z.infer<typeof this.ReadSheetDataSchema>): Promise<CallToolResult> {
    const { fileAbsolutePath, sheetName, range } = args;
    const workbook = await this.readWorkbook(fileAbsolutePath);
    const worksheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
    if (!worksheet) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${sheetName} not found`);
    }

    const maxChankCellsSize = 5000;

    let [startCol, startRow, endCol, endRow] = range ? this.parseRange(range) : [1, 1, worksheet.columnCount, worksheet.rowCount];

    const chunkRows = Math.max(1, Math.ceil(Math.abs(maxChankCellsSize / (endCol - startCol + 1))));
    const chunkEndRow = Math.min(endRow, startRow + chunkRows - 1);

    return {
      content: [{
        type: 'text',
        mimeType: 'text/html',
        text: this.createHtmlTable(worksheet, startCol, startRow, endCol, chunkEndRow)
      }]
    };
  }

  private createHtmlTable(worksheet: ExcelJS.Worksheet, startCol: number, startRow: number, endCol: number, endRow: number): string {
    const responseRange = `${this.columnNumberToLetter(startCol)}${startRow}:${this.columnNumberToLetter(endCol)}${endRow}`;
    const fullRange = `${this.columnNumberToLetter(1)}${1}:${this.columnNumberToLetter(worksheet.columnCount)}${worksheet.rowCount}`;
    let tableHtml = '<table>\n';
    tableHtml += `<tr><th>[${worksheet.name}] Current data range: ${responseRange}, Full data range: ${fullRange}</th>`;
    // 列アドレスを出力
    for (let col = startCol; col <= endCol; col++) {
      tableHtml += `<th>${this.columnNumberToLetter(col)}</th>`;
    }
    tableHtml += '</tr>\n';
    for (let row = startRow; row <= endRow; row++) {
      tableHtml += '<tr>';
      // 行アドレスを出力
      tableHtml += `<td>${row}</td>`;
      for (let col = startCol; col <= endCol; col++) {
        const cell = worksheet.getCell(row, col);
        const cellValue = cell.value ? cell.text.replaceAll('\n', '<br>') : '';
        tableHtml += `<td>${cellValue}</td>`;
      }
      tableHtml += '</tr>\n';
    }
    tableHtml += '</table>';
    return tableHtml;
  }

  /**
   * Write data to the Excel sheet
   * @returns Success message
   */
  private async handleWriteSheetData(args: z.infer<typeof this.WriteSheetDataSchema>): Promise<CallToolResult> {
    const { fileAbsolutePath, sheetName, range, data } = args;

    const workbook = await this.readWorkbook(fileAbsolutePath);
    const worksheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];

    if (!worksheet) {
      throw new McpError(ErrorCode.InvalidParams, `Sheet ${sheetName} not found`);
    }
    const dataColumnLength = Math.max(...data.map(row => row.length));

    // 範囲が指定されていない場合は、デフォルトで A1 から開始
    let startCol = 1, startRow = 1;
    let endCol: number, endRow: number;

    if (range) {
      [startCol, startRow, endCol, endRow] = this.parseRange(range);

      // データサイズと範囲サイズの整合性チェック
      const rangeRowCount = endRow - startRow + 1;
      const rangeColCount = endCol - startCol + 1;

      if (data.length != rangeRowCount) {
        throw new McpError(
          ErrorCode.InvalidParams,
          `Number of rows [${data.length}] of 'data' argument is not equal to the number of rows of specified range [${rangeRowCount}]`
        );
      }
      if (dataColumnLength != rangeColCount) {
        throw new McpError(
          ErrorCode.InvalidParams,
          `Number of columns [${dataColumnLength}] of 'data' argument is not equal to the number of columns of specified range [${rangeColCount}]`
        );
      }
    } else {
      // 範囲が指定されていない場合は、データサイズに基づいて範囲を決定
      endRow = startRow + data.length - 1;
      endCol = startCol + dataColumnLength - 1;
    }

    // 指定範囲にデータを書き込み
    for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
      const row = data[rowIndex];
      for (let columnIndex = 0; columnIndex < row.length; columnIndex++) {
        worksheet.getCell(startRow + rowIndex, startCol + columnIndex).value = row[columnIndex];
      }
    }
    await workbook.xlsx.writeFile(fileAbsolutePath);
    return {
      content: [{
        type: 'text',
        text: 'File saved successfully'
      }]
    };
  }

  /**
   * Reads Excel workbook from file
   * @param fileAbsolutePath Absolute path to the Excel file
   * @returns ExcelJS.Workbook
   */
  private async readWorkbook(fileAbsolutePath: string): Promise<ExcelJS.Workbook> {
    const workbook = this.workbookCache.get(fileAbsolutePath);
    if (workbook) {
      return workbook;
    } else {
      if (!fs.existsSync(fileAbsolutePath)) {
        throw new McpError(ErrorCode.InvalidParams, `File [${fileAbsolutePath}] not found`);
      }
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      this.workbookCache.set(fileAbsolutePath, workbook);
      return workbook;
    }
  }

  /**
   * Parses Excel range address into numeric coordinates
   * @param range Excel range string (e.g. "A1:C10")
   * @returns Array containing [startCol, startRow, endCol, endRow]
   * @throws McpError if range format is invalid
   */
  private parseRange(range: string): [number, number, number, number] {
    const match = range.match(/([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)/);
    if (!match) {
      throw new McpError(ErrorCode.InvalidParams, 'Invalid range address format. Expected format like "A1:C10"');
    }
    const [_, startColMatch, startRowMatch, endColMatch, endRowMatch] = match;

    const startCol = this.columnLetterToNumber(startColMatch);
    const startRow = parseInt(startRowMatch, 10);
    const endCol = this.columnLetterToNumber(endColMatch);
    const endRow = parseInt(endRowMatch, 10);

    return [startCol, startRow, endCol, endRow];
  }

  /**
   * Converts Excel column letters to numeric index
   * @param letters Column letters (e.g. "A", "B", "AA")
   * @returns Numeric column index (1-based)
   */
  private columnLetterToNumber(letters: string): number {
    let column = 0;
    letters = letters.toUpperCase();
    for (let i = 0; i < letters.length; i++) {
      column = column * 26 + (letters.charCodeAt(i) - 64);
    }
    return column;
  }

  /**
   * Converts Excel column numeric index to column letters
   * @param num Numeric column index (1-based)
   * @returns Column letters (e.g. 2 -> "B", 27 -> "AA")
   */
  private columnNumberToLetter(num: number): string {
    let letters = '';
    while (num > 0) {
      const remainder = (num - 1) % 26;
      letters = String.fromCharCode(65 + remainder) + letters;
      num = Math.floor((num - 1) / 26);
    }
    return letters;
  }
};

const server = new ExcelMcpServer();
server.start().catch(console.error);
