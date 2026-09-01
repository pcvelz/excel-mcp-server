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

type ExcelMoveSheetArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Index            int    `zog:"index"`
}

var excelMoveSheetArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"index":            z.Int().GTE(0).Required(),
})

func AddExcelMoveSheetTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_move_sheet",
		mcp.WithDescription("Move a sheet to another position in the Excel file's tab order. Content and formatting are untouched."),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Name of the sheet to move"),
		),
		mcp.WithNumber("index",
			mcp.Required(),
			mcp.Description("Zero-based target position. 0 makes the sheet the first tab, which is the one shown when the workbook is opened."),
		),
	), handleMoveSheet)
}

func handleMoveSheet(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelMoveSheetArguments{}
	if issues := excelMoveSheetArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return moveSheet(args.FileAbsolutePath, args.SheetName, args.Index)
}

func moveSheet(fileAbsolutePath string, sheetName string, index int) (*mcp.CallToolResult, error) {
	workbook, release, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer release()

	if err := workbook.MoveSheet(sheetName, index); err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Sheet [%s] moved to index %d.\n", html.EscapeString(sheetName), index)
	result += formatSheetOrder(workbook)
	return mcp.NewToolResultText(result), nil
}
