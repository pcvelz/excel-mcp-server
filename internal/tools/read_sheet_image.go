package tools

import (
	"context"
	"fmt"
	"strings"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ReadSheetImageArguments struct {
	FileAbsolutePath  string   `zog:"fileAbsolutePath"`
	SheetName         string   `zog:"sheetName"`
	Range             string   `zog:"range"`
	KnownPagingRanges []string `zog:"knownPagingRanges"`
}

var readSheetImageArgumentsSchema = z.Struct(z.Schema{
	"fileAbsolutePath":  z.String().Required(),
	"sheetName":         z.String().Required(),
	"range":             z.String(),
	"knownPagingRanges": z.Slice(z.String()),
})

func AddReadSheetImageTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("read_sheet_image",
		mcp.WithDescription("[Windows only] Read data as an image from the Excel sheet with pagination."),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("range",
			mcp.Description("Range of cells to read in the Excel sheet (e.g., \"A1:C10\"). [default: first paging range]"),
		),
		mcp.WithArray("knownPagingRanges",
			mcp.Description("List of already read paging ranges"),
			mcp.Items(map[string]any{
				"type": "string",
			}),
		),
	), handleReadSheetImage)
}

func handleReadSheetImage(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ReadSheetImageArguments{}
	issues := readSheetImageArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return readSheetImage(args.FileAbsolutePath, args.SheetName, args.Range, args.KnownPagingRanges)
}

func readSheetImage(fileAbsolutePath string, sheetName string, rangeStr string, knownPagingRanges []string) (*mcp.CallToolResult, error) {
	workbook, releaseWorkbook, err := excel.NewExcelOle(fileAbsolutePath)
	defer releaseWorkbook()
	if err != nil {
		workbook, releaseWorkbook, err = excel.NewExcelOleWithNewObject(fileAbsolutePath)
		defer releaseWorkbook()
		if err != nil {
			return imcp.NewToolResultInvalidArgumentError(fmt.Errorf("failed to open workbook: %w", err).Error()), nil
		}
	}
	// シート取得
	var worksheet excel.Worksheet
	if sheetName == "" {
		// シート名未指定時は "Sheet1" を仮定
		worksheet, err = workbook.FindSheet("Sheet1")
		if err != nil {
			return imcp.NewToolResultInvalidArgumentError("sheet not found"), nil
		}
		sheetName, _ = worksheet.Name()
	} else {
		worksheet, err = workbook.FindSheet(sheetName)
		if err != nil {
			return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("sheet %s not found", sheetName)), nil
		}
	}

	pagingStrategy, err := worksheet.GetPagingStrategy(5000)
	if err != nil {
		return nil, err
	}
	pagingService := excel.NewPagingRangeService(pagingStrategy)

	allRanges := pagingService.GetPagingRanges()
	if len(allRanges) == 0 {
		return imcp.NewToolResultInvalidArgumentError("no range available to read"), nil
	}

	var currentRange string
	if rangeStr == "" {
		// range が指定されていない場合は最初の Range を使用
		currentRange = allRanges[0]
	} else {
		// range が指定されている場合は指定された範囲を使用
		currentRange = rangeStr
	}
	remainingRanges := pagingService.FilterRemainingPagingRanges(allRanges, append(knownPagingRanges, currentRange))

	base64image, err := worksheet.CapturePicture(currentRange)
	if err != nil {
		return nil, fmt.Errorf("failed to copy range to image: %w", err)
	}

	text := "# Metadata\n"
	text += fmt.Sprintf("- sheet name: %s\n", sheetName)
	text += fmt.Sprintf("- read range: %s\n", currentRange)
	text += "# Notice\n"
	if len(remainingRanges) > 0 {
		text += "This sheet has more some ranges.\n"
		text += "To read the next range, you should specify 'range' and 'knownPagingRanges' arguments as follows.\n"
		text += fmt.Sprintf("`{ \"range\": \"%s\", \"knownPagingRanges\": [%s] }`", remainingRanges[0], "\""+strings.Join(append(knownPagingRanges, currentRange), "\", \"")+"\"")
	} else {
		text += "All ranges have been read.\n"
	}

	// 結果を返却
	return mcp.NewToolResultImage(
		text,
		base64image,
		"image/png",
	), nil
}
