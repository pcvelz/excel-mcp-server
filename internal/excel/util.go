package excel

import (
	"fmt"
	"math"
	"regexp"

	"github.com/xuri/excelize/v2"
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
