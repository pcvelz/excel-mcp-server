package tools

import (
	"context"
	"fmt"
	"html"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	excel "github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ExcelInsertRowsArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	BeforeRow        int    `zog:"beforeRow"`
	Count            int    `zog:"count"`
}

var excelInsertRowsArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"beforeRow":        z.Int().GTE(1).Required(),
	"count":            z.Int().GTE(1).Default(1),
})

func AddExcelInsertRowsTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_insert_rows",
		mcp.WithDescription("Insert empty rows into the Excel sheet and shift the rows at and below the insertion point down. Merged cells, conditional formatting and data validation move with the rows, and formulas are adjusted the way Excel adjusts them."),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithNumber("beforeRow",
			mcp.Required(),
			mcp.Description("One-based row number to insert before. The new rows take this position and the existing row moves down."),
		),
		mcp.WithNumber("count",
			mcp.Description("Number of rows to insert. Defaults to 1."),
		),
	), handleInsertRows)
}

func handleInsertRows(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelInsertRowsArguments{}
	if issues := excelInsertRowsArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return insertRows(args.FileAbsolutePath, args.SheetName, args.BeforeRow, args.Count)
}

func insertRows(fileAbsolutePath string, sheetName string, beforeRow int, count int) (*mcp.CallToolResult, error) {
	workbook, release, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer release()

	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	defer worksheet.Release()

	if err := worksheet.InsertRows(beforeRow, count); err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Inserted %d empty row(s) before row %d in sheet [%s].\n", count, beforeRow, html.EscapeString(sheetName))
	result += formatSheetRules(worksheet)
	// Excel itself copies the formatting of the row above into inserted rows;
	// excelize leaves them bare.
	if workbook.GetBackendName() == "excelize" {
		result += "\nInserted rows carry no formatting. Use excel_format_range if the new rows should match their neighbours.\n"
	}
	return mcp.NewToolResultText(result), nil
}
