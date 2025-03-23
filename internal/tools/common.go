package tools

import (
	"fmt"
	"regexp"
	"strings"

	"github.com/devlights/goxcel"
	"github.com/go-ole/go-ole/oleutil"
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

func CreateHTMLTableOfValues(workbook *excelize.File, sheetName string, startCol int, startRow int, endCol int, endRow int) (*string, error) {
	return createHTMLTable(workbook, sheetName, startCol, startRow, endCol, endRow, func(sheetName string, cellRange string) (string, error) {
		return workbook.GetCellValue(sheetName, cellRange)
	})
}

func CreateHTMLTableOfFormula(workbook *excelize.File, sheetName string, startCol int, startRow int, endCol int, endRow int) (*string, error) {
	return createHTMLTable(workbook, sheetName, startCol, startRow, endCol, endRow, func(sheetName string, cellRange string) (string, error) {
		formula, err := workbook.GetCellFormula(sheetName, cellRange)
		if err != nil {
			return "", err
		}
		if formula == "" {
			// fallback
			return workbook.GetCellValue(sheetName, cellRange)
		}
		if !strings.HasPrefix(formula, "=") {
			formula = "=" + formula
		}
		return formula, nil
	})
}

// CreateHTMLTable creates a table data in HTML format
func createHTMLTable(workbook *excelize.File, sheetName string, startCol int, startRow int, endCol int, endRow int, extractor func(sheetName string, cellRange string) (string, error)) (*string, error) {
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
			value, _ := extractor(sheetName, axis)
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
