package tools

import (
	"context"
	"fmt"
	"html"
	"strings"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	excel "github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ExcelDeleteRowsArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	StartRow         int    `zog:"startRow"`
	EndRow           int    `zog:"endRow"`
}

var excelDeleteRowsArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"startRow":         z.Int().GTE(1).Required(),
	"endRow":           z.Int().GTE(1).Required(),
})

func AddExcelDeleteRowsTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_delete_rows",
		mcp.WithDescription("Delete rows from the Excel sheet and shift the rows below them up. Formatting of the remaining rows, merged cells, column widths, conditional formatting and data validation are preserved, and formulas are adjusted the way Excel adjusts them: references into the deleted rows become #REF!."),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithNumber("startRow",
			mcp.Required(),
			mcp.Description("First row to delete, one-based and inclusive"),
		),
		mcp.WithNumber("endRow",
			mcp.Required(),
			mcp.Description("Last row to delete, one-based and inclusive. Use the same value as startRow to delete a single row."),
		),
	), handleDeleteRows)
}

func handleDeleteRows(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelDeleteRowsArguments{}
	if issues := excelDeleteRowsArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return deleteRows(args.FileAbsolutePath, args.SheetName, args.StartRow, args.EndRow)
}

func deleteRows(fileAbsolutePath string, sheetName string, startRow int, endRow int) (*mcp.CallToolResult, error) {
	if endRow < startRow {
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("endRow (%d) must not be smaller than startRow (%d)", endRow, startRow)), nil
	}

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

	if err := worksheet.DeleteRows(startRow, endRow); err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	count := endRow - startRow + 1
	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Deleted %d row(s) %d-%d from sheet [%s].\n", count, startRow, endRow, html.EscapeString(sheetName))
	result += formatSheetRules(worksheet)
	return mcp.NewToolResultText(result), nil
}

// formatSheetRules reports the used range plus the sheet-level rules that a row
// operation could damage, so the caller can confirm they survived rather than
// having to trust that they did.
func formatSheetRules(worksheet excel.Worksheet) string {
	result := ""
	if usedRange, err := worksheet.GetDimention(); err == nil {
		result += fmt.Sprintf("Used range is now: %s\n", usedRange)
	}
	if merged, err := worksheet.GetMergedCells(); err == nil && len(merged) > 0 {
		result += fmt.Sprintf("Merged cells: %s\n", html.EscapeString(strings.Join(merged, ", ")))
	}
	if formats, err := worksheet.GetConditionalFormatRanges(); err == nil && len(formats) > 0 {
		result += fmt.Sprintf("Conditional formatting on: %s\n", html.EscapeString(strings.Join(formats, ", ")))
	}
	if validations, err := worksheet.GetDataValidationRanges(); err == nil && len(validations) > 0 {
		result += fmt.Sprintf("Data validation on: %s\n", html.EscapeString(strings.Join(validations, ", ")))
	}
	return result
}
