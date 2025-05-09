package tools

import (
	"context"
	"encoding/json"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ExcelDescribeSheetsArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
}

var excelDescribeSheetsArgumentsSchema = z.Struct(z.Schema{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
})

func AddExcelDescribeSheetsTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_describe_sheets",
		mcp.WithDescription("List all sheet names in an Excel file"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
	), handleDescribeSheets)
}

func handleDescribeSheets(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelDescribeSheetsArguments{}
	issues := excelDescribeSheetsArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return describeSheets(args.FileAbsolutePath)
}

func describeSheets(fileAbsolutePath string) (*mcp.CallToolResult, error) {
	workbook, release, err := excel.OpenFile(fileAbsolutePath)
	defer release()
	if err != nil {
		return nil, err
	}

	sheetList, err := workbook.GetSheetNames()
	if err != nil {
		return nil, err
	}
	jsonBytes, err := json.MarshalIndent(sheetList, "", "  ")
	if err != nil {
		return nil, err
	}

	return mcp.NewToolResultText(string(jsonBytes)), nil
}
