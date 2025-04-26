package tools

import (
	"context"
	"fmt"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
	"github.com/xuri/excelize/v2"
)

type WriteSheetDataArguments struct {
	FileAbsolutePath string     `zog:"fileAbsolutePath"`
	SheetName        string     `zog:"sheetName"`
	Range            string     `zog:"range"`
	Data             [][]string `zog:"data"`
}

var writeSheetDataArgumentsSchema = z.Struct(z.Schema{
	"fileAbsolutePath": z.String().Required(),
	"sheetName":        z.String().Required(),
	"range":            z.String().Required(),
	"data":             z.Slice(z.Slice(z.String())).Required(),
})

func AddWriteSheetDataTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("write_sheet_data",
		mcp.WithDescription("Write data to the Excel sheet"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("range",
			mcp.Required(),
			mcp.Description("Range of cells in the Excel sheet (e.g., \"A1:C10\")"),
		),
		mcp.WithArray("data",
			mcp.Required(),
			mcp.Description("Data to write to the Excel sheet"),
			mcp.Items(map[string]any{
				"type": "array",
				"items": map[string]any{
					"anyOf": []any{
						map[string]any{
							"type": "string",
						},
						map[string]any{
							"type": "number",
						},
						map[string]any{
							"type": "boolean",
						},
						map[string]any{
							"type": "null",
						},
					},
				},
			}),
		),
	), handleWriteSheetData)
}

func handleWriteSheetData(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := WriteSheetDataArguments{}
	issues := writeSheetDataArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}

	// zog が any type のスキーマをサポートしていないため、自力で実装
	dataArg, ok := request.Params.Arguments["data"].([]any)
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("data must be a 2D array"), nil
	}
	data := make([][]any, len(dataArg))
	for i, v := range dataArg {
		value, ok := v.([]any)
		if !ok {
			return imcp.NewToolResultInvalidArgumentError("data must be a 2D array"), nil
		}
		data[i] = value
	}

	return writeSheetData(args.FileAbsolutePath, args.SheetName, args.Range, data)
}

func writeSheetData(fileAbsolutePath string, sheetName string, rangeStr string, data [][]any) (*mcp.CallToolResult, error) {
	workbook, closeFn, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer closeFn()

	// シートの取得
	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}

	startCol, startRow, endCol, endRow, err := excel.ParseRange(rangeStr)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}

	// データの整合性チェック
	rangeRowSize := endRow - startRow + 1
	if len(data) != rangeRowSize {
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("number of rows in data (%d) does not match range size (%d)", len(data), rangeRowSize)), nil
	}

	// データの書き込み
	for i, row := range data {
		rangeColumnSize := endCol - startCol + 1
		if len(row) != rangeColumnSize {
			return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("number of columns in row %d (%d) does not match range size (%d)", i, len(row), rangeColumnSize)), nil
		}
		for j, cellValue := range row {
			cell, err := excelize.CoordinatesToCellName(startCol+j, startRow+i)
			if err != nil {
				return nil, err
			}
			err = worksheet.SetValue(cell, cellValue)
			if err != nil {
				return nil, err
			}
		}
	}

	if err := workbook.Save(); err != nil {
		return nil, err
	}

	// HTMLテーブルの生成
	table, err := CreateHTMLTableOfFormula(worksheet, startCol, startRow, endCol, endRow)
	if err != nil {
		return nil, err
	}
	html := "<h2>Sheet Data</h2>\n"
	html += *table + "\n"
	html += "<h2>Metadata</h2>\n"
	html += "<ul>\n"
	html += fmt.Sprintf("<li>sheet name: %s</li>\n", sheetName)
	html += fmt.Sprintf("<li>read range: %s</li>\n", rangeStr)
	html += "</ul>\n"
	html += "<h2>Notice</h2>\n"
	html += "<p>Values wrote successfully.</p>\n"

	return mcp.NewToolResultText(html), nil
}
