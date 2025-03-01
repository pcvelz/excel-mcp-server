package tools

import (
	"context"
	"fmt"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
	"github.com/xuri/excelize/v2"
)

type ReadSheetDataArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Range            string `zog:"range"`
}

var readSheetDataArgumentsSchema = z.Struct(z.Schema{
	"fileAbsolutePath": z.String().Required(),
	"sheetName":        z.String().Required(),
	"range":            z.String(),
})

func AddReadSheetDataTool(server *server.MCPServer) {
  server.AddTool(mcp.NewTool("read_sheet_data",
    mcp.WithDescription("Read data from the Excel sheet."+
      "If there is a large number of data, it reads a part of data."+
      "To read more data, adjust range parameter and make requests again."),
    mcp.WithString("fileAbsolutePath",
      mcp.Required(),
      mcp.Description("Absolute path to the Excel file"),
    ),
    mcp.WithString("sheetName",
      mcp.Required(),
      mcp.Description("Sheet name in the Excel file"),
    ),
    mcp.WithString("range",
      mcp.Description("Range of cells in the Excel sheet (e.g., \"A1:C10\")"),
    ),
  ), handleReadSheetData)
}

// HandleReadSheetData handles read_sheet_data tool request
func handleReadSheetData(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ReadSheetDataArguments{}
	issues := readSheetDataArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	filePath := args.FileAbsolutePath
	workbook, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	defer workbook.Close()

	sheetName := args.SheetName
	if sheetName == "" {
		sheetName = workbook.GetSheetList()[0]
	}
	index, _ := workbook.GetSheetIndex(sheetName)
	if index == -1 {
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("sheet %s not found", sheetName)), nil
	}
	rangeStr := args.Range

	dimension, err := workbook.GetSheetDimension(sheetName)
	if err != nil {
		return nil, err
	}
	var startCol, startRow, endCol, endRow int
	if rangeStr == "" {
		startCol, startRow, endCol, endRow, err = ParseRange(dimension)
	} else {
		startCol, startRow, endCol, endRow, err = ParseRange(rangeStr)
	}
	if err != nil {
		return nil, err
	}

	// データ量が多い場合は制限する
	maxChunkCells := 5000
	chunkRows := max(1, maxChunkCells/(endCol-startCol+1))
	endRow = min(endRow, startRow+chunkRows-1)

	// HTML テーブルの生成
	table, err := CreateHTMLTable(workbook, sheetName, startCol, startRow, endCol, endRow)
	if err != nil {
		return nil, err
	}

	return mcp.NewToolResultText(*table), nil
}
