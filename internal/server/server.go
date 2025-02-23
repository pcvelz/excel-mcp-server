package server

import (
	"context"
	"fmt"
	"regexp"
	"strings"

	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
	"github.com/xuri/excelize/v2"
)

type ExcelServer struct {
	server *server.MCPServer
}

func New() *ExcelServer {
	s := &ExcelServer{}
	s.server = server.NewMCPServer(
		"excel-mcp-server",
		"0.1.5",
	)
	// ツールの登録
	s.server.AddTool(mcp.NewTool("read_sheet_names",
		mcp.WithDescription("List all sheet names in an Excel file"),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
	), s.handleReadSheetNames)
	s.server.AddTool(mcp.NewTool("read_sheet_data",
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
	), s.handleReadSheetData)
	s.server.AddTool(mcp.NewTool("write_sheet_data",
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
		imcp.WithArray("data",
			mcp.Required(),
			mcp.Description("Data to write to the Excel sheet"),
		),
	), s.handleWriteSheetData)
	return s
}

func (s *ExcelServer) Start() error {
	return server.ServeStdio(s.server)
}

// List all sheet names in an Excel file
// Response: Sheet names
func (s *ExcelServer) handleReadSheetNames(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	filePath, ok := request.Params.Arguments["fileAbsolutePath"].(string)
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("fileAbsolutePath must be a string"), nil
	}
	workbook, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	defer workbook.Close()

	sheetList := workbook.GetSheetList()
	var sheetNames []any
	for _, name := range sheetList {
		sheetNames = append(sheetNames, mcp.NewTextContent(name))
	}

	return &mcp.CallToolResult{
		Content: sheetNames,
	}, nil
}

// Read data from the Excel sheet
// Response: Spreadsheet data in HTML table format
func (s *ExcelServer) handleReadSheetData(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	filePath, ok := request.Params.Arguments["fileAbsolutePath"].(string)
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("fileAbsolutePath must be a string"), nil
	}
	workbook, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	defer workbook.Close()

	sheetName, _ := request.Params.Arguments["sheetName"].(string)
	if sheetName == "" {
		sheetName = workbook.GetSheetList()[0]
	}
	index, _ := workbook.GetSheetIndex(sheetName)
	if index == -1 {
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("sheet %s not found", sheetName)), nil
	}
	rangeStr, ok := request.Params.Arguments["range"].(string)
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("range must be a string"), nil
	}

	dimension, err := workbook.GetSheetDimension(sheetName)
	if err != nil {
		return nil, err
	}
	var startCol, startRow, endCol, endRow int
	if rangeStr == "" {
		startCol, startRow, endCol, endRow, err = s.parseRange(dimension)
	} else {
		startCol, startRow, endCol, endRow, err = s.parseRange(rangeStr)
	}
	if err != nil {
		return nil, err
	}

	// データ量が多い場合は制限する
	maxChunkCells := 5000
	chunkRows := max(1, maxChunkCells/(endCol-startCol+1))
	endRow = min(endRow, startRow+chunkRows-1)

	// HTML テーブルの生成
	table, err := s.createHTMLTable(workbook, sheetName, startCol, startRow, endCol, endRow)
	if err != nil {
		return nil, err
	}

	return mcp.NewToolResultText(*table), nil
}

// Write data to the Excel sheet
// Response: Success message
func (s *ExcelServer) handleWriteSheetData(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	filePath, ok := request.Params.Arguments["fileAbsolutePath"].(string)
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("fileAbsolutePath must be a string"), nil
	}
	sheetName, ok := request.Params.Arguments["sheetName"].(string)
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("sheetName must be a string"), nil
	}
	rangeStr, ok := request.Params.Arguments["range"].(string)
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("range must be a string"), nil
	}
	data, ok := request.Params.Arguments["data"].([]any)
	if !ok {
		return imcp.NewToolResultInvalidArgumentError("data must be a 2D array"), nil
	}
	rangeData := make([][]any, len(data))
	for i, v := range data {
		value, ok := v.([]any)
		if !ok {
			return imcp.NewToolResultInvalidArgumentError("data must be a 2D array"), nil
		}
		rangeData[i] = value
	}

	workbook, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	defer workbook.Close()

	// シートの取得
	index, _ := workbook.GetSheetIndex(sheetName)
	if index == -1 {
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("sheet %s not found", sheetName)), nil
	}

	startCol, startRow, endCol, endRow, err := s.parseRange(rangeStr)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}

	// データの整合性チェック
	rangeRowSize := endRow - startRow + 1
	if len(rangeData) != rangeRowSize {
		return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("number of rows in data (%d) does not match range size (%d)", len(data), rangeRowSize)), nil
	}

	// データの書き込み
	for i, row := range rangeData {
		rangeColumnSize := endCol - startCol + 1
		if len(row) != rangeColumnSize {
			return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("number of columns in row %d (%d) does not match range size (%d)", i, len(row), rangeColumnSize)), nil
		}
		for j, cellValue := range row {
			cell, err := excelize.CoordinatesToCellName(startCol+j, startRow+i)
			if err != nil {
				return nil, err
			}
			err = workbook.SetCellValue(sheetName, cell, cellValue)
			if err != nil {
				return nil, err
			}
		}
	}

	if err := workbook.Save(); err != nil {
		return nil, err
	}

	return mcp.NewToolResultText("File saved successfully"), nil
}

// parseRange parses Excel's range string (e.g. A1:C10)
func (s *ExcelServer) parseRange(rangeStr string) (int, int, int, int, error) {
	re := regexp.MustCompile(`(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)`)
	matches := re.FindStringSubmatch(rangeStr)
	if matches == nil {
		return 0, 0, 0, 0, fmt.Errorf("invalid range format: %s", rangeStr)
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(matches[1])
	if err != nil {
		return 0, 0, 0, 0, err
	}
	endCol, endRow, err := excelize.CellNameToCoordinates(matches[2])
	if err != nil {
		return 0, 0, 0, 0, err
	}
	return startCol, startRow, endCol, endRow, nil
}

// createHTMLTable creates a table dada in HTML format
func (s *ExcelServer) createHTMLTable(workbook *excelize.File, sheetName string, startCol int, startRow int, endCol int, endRow int) (*string, error) {
	var table string
	table += "<table>\n"

	// ヘッダー行（範囲情報）
	startCell, err := excelize.CoordinatesToCellName(startCol, startRow)
	if err != nil {
		return nil, err
	}
	endCell, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		return nil, err
	}
	responseRange := fmt.Sprintf("%s:%s", startCell, endCell)
	fullRange, err := workbook.GetSheetDimension(sheetName)
	if err != nil {
		return nil, err
	}
	table += fmt.Sprintf("<tr><th>[%s] Current data range: %s, Full data range: %s</th>",
		sheetName, responseRange, fullRange)

	// 列アドレスの出力
	for col := startCol; col <= endCol; col++ {
		name, _ := excelize.ColumnNumberToName(col)
		table += fmt.Sprintf("<th>%s</th>", name)
	}
	table += "</tr>\n"

	// データの出力
	for row := startRow; row <= endRow; row++ {
		table += "<tr>"
		// 行アドレスを出力
		table += fmt.Sprintf("<td>%d</td>", row)

		for col := startCol; col <= endCol; col++ {
			axis, _ := excelize.CoordinatesToCellName(col, row)
			value, _ := workbook.GetCellValue(sheetName, axis)
			table += fmt.Sprintf("<td>%s</td>", strings.ReplaceAll(value, "\n", "<br>"))
		}
		table += "</tr>\n"
	}

	table += "</table>"
	return &table, nil
}
