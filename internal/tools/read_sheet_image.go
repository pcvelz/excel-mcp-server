package tools

import (
	"bufio"
	"bytes"
	"context"
	"encoding/base64"
	"fmt"
	"strings"

	z "github.com/Oudwins/zog"
	"github.com/devlights/goxcel"
	"github.com/devlights/goxcel/constants"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ReadSheetImageArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	Range            string `zog:"range"`
}

var readSheetImageArgumentsSchema = z.Struct(z.Schema{
	"fileAbsolutePath": z.String().Required(),
	"sheetName":        z.String().Required(),
	"range":            z.String(),
})

func AddReadSheetImageTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("read_sheet_image",
		mcp.WithDescription("[Windows only] Read data as an image from the Excel sheet."+
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
			mcp.Description("Range of cells in the Excel sheet (e.g., \"A1:C10\")."+
				"If not specified, it try to read the entire sheet."),
		),
	), handleReadSheetImage)
}

func handleReadSheetImage(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ReadSheetImageArguments{}
	issues := readSheetImageArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}

	quit := goxcel.MustInitGoxcel()
	defer quit()

	excel, release := goxcel.MustNewGoxcel()
	defer release()

	// ワークブックを開く
	workbooks := excel.MustWorkbooks()
	workbook, releaseWorkbook, err := workbooks.Open(args.FileAbsolutePath)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(fmt.Errorf("failed to open workbook: %w", err).Error()), nil
	}
	defer releaseWorkbook()

	// ワークシートを取得
	sheets := workbook.MustWorkSheets()
	sheet, _ := sheets.Walk(func(sheet *goxcel.Worksheet, i int) error {
		name, err := sheet.Name()
		if err != nil {
			return nil
		}
		if name == args.SheetName {
			return fmt.Errorf("found")
		}
		return nil
	})
	if sheet == nil {
		return nil, fmt.Errorf("sheet not found: %s", args.SheetName)
	}
	sheetName, err := sheet.Name()
	if err != nil {
		return nil, fmt.Errorf("failed to get sheet name: %w", err)
	}

	// range を処理
	xlUsedRange, err := sheet.UsedRange()
	if err != nil {
		return nil, fmt.Errorf("failed to get used range: %w", err)
	}
	var xlRange *goxcel.XlRange
	if args.Range == "" {
		// range が指定されていない場合は UsedRange を使用
		xlRange = xlUsedRange
	} else {
		// range が指定されている場合は指定された範囲を使用
		startCol, startRow, endCol, endRow, err := ParseRange(args.Range)
		if err != nil {
			return imcp.NewToolResultInvalidArgumentError(fmt.Errorf("failed to parse range: %w", err).Error()), nil
		}
		// 必要な部分だけを選択するようにする
		xlRange, err = sheet.Range(startRow, startCol, endRow, endCol)
		if err != nil {
			return nil, fmt.Errorf("failed to create range: %w", err)
		}
	}
	usedRangeAddress, err := fetchRangeAddress(xlUsedRange)
	if err != nil {
		return nil, fmt.Errorf("failed to get address: %w", err)
	}
	rangeAddress, err := fetchRangeAddress(xlRange)
	if err != nil {
		return nil, fmt.Errorf("failed to get address: %w", err)
	}
	// 画像の取得用バッファの準備
	buf := new(bytes.Buffer)
	bufWriter := bufio.NewWriter(buf)
	err = xlRange.CopyPictureToFile(bufWriter, constants.XlScreen, constants.XlBitmap)
	if err != nil {
		return nil, fmt.Errorf("failed to copy range to image: %w", err)
	}
	err = bufWriter.Flush()
	if err != nil {
		return nil, fmt.Errorf("failed to flush buffer: %w", err)
	}
	// base64 エンコード
	base64Image := base64.StdEncoding.EncodeToString(buf.Bytes())

	// 結果を返却
	return mcp.NewToolResultImage(
		fmt.Sprintf("[%s] Current data range: %s, Full data range: %s", sheetName, rangeAddress, usedRangeAddress),
		base64Image,
		"image/png",
	), nil
}

func fetchRangeAddress(XlRange *goxcel.XlRange) (string, error) {
	address, err := oleutil.GetProperty(XlRange.ComObject(), "Address")
	if err != nil {
		return "", err
	}
	return strings.ReplaceAll(address.ToString(), "$", ""), nil
}
