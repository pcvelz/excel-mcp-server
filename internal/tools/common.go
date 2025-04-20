package tools

import (
	"fmt"
	"math"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/devlights/goxcel"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/xuri/excelize/v2"

  "github.com/negokaz/excel-mcp-server/internal/excel"
)

// parseRange parses Excel's range string (e.g. A1:C10)
func ParseRange(rangeStr string) (int, int, int, int, error) {
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

// GetSheetDimensionByIterators returns the actual data range in A1 format for the sheet.
// Uses Cols/Rows iterators to calculate the exact range.
func GetSheetDimensionByIterators(f *excelize.File, sheetName string) (string, error) {
	minRow, maxRow := math.MaxInt, 0
	minCol, maxCol := math.MaxInt, 0

	cols, err := f.Cols(sheetName)
	if err != nil {
		return "", fmt.Errorf("failed to get columns iterator: %w", err)
	}

	colIdx := 1
	for cols.Next() {
		rowData, err := cols.Rows()
		if err != nil {
			return "", fmt.Errorf("failed to get row data: %w", err)
		}

		currentColumnHasData := false
		for rowIdx, val := range rowData {
			if val != "" {
				currentColumnHasData = true
				rowNum := rowIdx + 1
				if rowNum < minRow {
					minRow = rowNum
				}
				if rowNum > maxRow {
					maxRow = rowNum
				}
			}
		}
		if currentColumnHasData {
			if colIdx < minCol {
				minCol = colIdx
			}
			if colIdx > maxCol {
				maxCol = colIdx
			}
		}
		colIdx++
	}

	// If no data exists
	if maxRow == 0 || maxCol == 0 {
		minCol, maxCol, minRow, maxRow = 1, 1, 1, 1
	}

	startCell, err := excelize.CoordinatesToCellName(minCol, minRow)
	if err != nil {
		return "", fmt.Errorf("failed to convert coordinates to cell name: %w", err)
	}
	endCell, err := excelize.CoordinatesToCellName(maxCol, maxRow)
	if err != nil {
		return "", fmt.Errorf("failed to convert coordinates to cell name: %w", err)
	}

	dimension := fmt.Sprintf("%s:%s", startCell, endCell)
	return dimension, nil
}

func CreateHTMLTableOfValues(worksheet excel.Worksheet, startCol int, startRow int, endCol int, endRow int) (*string, error) {
	return createHTMLTable(worksheet, startCol, startRow, endCol, endRow, func(cellRange string) (string, error) {
		return worksheet.GetValue(cellRange)
	})
}

func CreateHTMLTableOfFormula(worksheet excel.Worksheet, startCol int, startRow int, endCol int, endRow int) (*string, error) {
	return createHTMLTable(worksheet, startCol, startRow, endCol, endRow, func(cellRange string) (string, error) {
		formula, err := worksheet.GetFormula(cellRange)
		if err != nil {
			return "", err
		}
		if formula == "" {
			// fallback
			return worksheet.GetValue(cellRange)
		}
		if !strings.HasPrefix(formula, "=") {
			formula = "=" + formula
		}
		return formula, nil
	})
}

// CreateHTMLTable creates a table data in HTML format
func createHTMLTable(worksheet excel.Worksheet, startCol int, startRow int, endCol int, endRow int, extractor func(cellRange string) (string, error)) (*string, error) {
	table := "<table>\n<tr><th></th>"

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
		table += fmt.Sprintf("<th>%d</th>", row)

		for col := startCol; col <= endCol; col++ {
			axis, _ := excelize.CoordinatesToCellName(col, row)
			value, _ := extractor(axis)
			table += fmt.Sprintf("<td>%s</td>", strings.ReplaceAll(value, "\n", "<br>"))
		}
		table += "</tr>\n"
	}

	table += "</table>"
	return &table, nil
}

func FetchRangeAddress(r *goxcel.XlRange) (string, error) {
	address, err := oleutil.GetProperty(r.ComObject(), "Address")
	if err != nil {
		return "", err
	}
	return strings.ReplaceAll(address.ToString(), "$", ""), nil
}

// SaveExcelize saves the Excel file to the specified path.
// Excelize's Save method restricts the file path length to 207 characters,
// but since this limitation has been relaxed in some environments,
// we ignore this restriction.
// https://github.com/qax-os/excelize/blob/v2.9.0/file.go#L71-L73
func SaveExcelize(f *excelize.File) error {
	file, err := os.OpenFile(filepath.Clean(f.Path), os.O_WRONLY|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if err != nil {
		return err
	}
	defer file.Close()
	return f.Write(file)
}
