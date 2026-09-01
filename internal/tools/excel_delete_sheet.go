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

type ExcelDeleteSheetArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Force            bool   `zog:"force"`
}

var excelDeleteSheetArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"force":            z.Bool().Default(false),
})

func AddExcelDeleteSheetTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_delete_sheet",
		mcp.WithDescription("Delete a sheet from the Excel file. Refuses to delete the last remaining sheet, and refuses to delete a sheet that formulas, data validations, charts, tables or pivot tables elsewhere in the workbook still refer to unless force is set. Defined names pointing at the deleted sheet are removed."),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Name of the sheet to delete"),
		),
		mcp.WithBoolean("force",
			mcp.Description("Delete even when formulas, data validations, charts, tables or pivot tables elsewhere refer to this sheet, leaving those references broken. Defaults to false."),
		),
	), handleDeleteSheet)
}

func handleDeleteSheet(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelDeleteSheetArguments{}
	if issues := excelDeleteSheetArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return deleteSheet(args.FileAbsolutePath, args.SheetName, args.Force)
}

func deleteSheet(fileAbsolutePath string, sheetName string, force bool) (*mcp.CallToolResult, error) {
	workbook, release, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer release()

	warnings, err := workbook.DeleteSheet(sheetName, force)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Sheet [%s] deleted.\n", html.EscapeString(sheetName))
	result += formatSheetOrder(workbook)
	result += formatWarnings(warnings)
	return mcp.NewToolResultText(result), nil
}
