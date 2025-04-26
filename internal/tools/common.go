package tools

import (
	"fmt"
	"strings"
  "html"

	"github.com/xuri/excelize/v2"

	"github.com/negokaz/excel-mcp-server/internal/excel"
)

func CreateHTMLTableOfValues(worksheet excel.Worksheet, startCol int, startRow int, endCol int, endRow int) (*string, error) {
	return createHTMLTable(startCol, startRow, endCol, endRow, func(cellRange string) (string, error) {
		return worksheet.GetValue(cellRange)
	})
}

func CreateHTMLTableOfFormula(worksheet excel.Worksheet, startCol int, startRow int, endCol int, endRow int) (*string, error) {
	return createHTMLTable(startCol, startRow, endCol, endRow, func(cellRange string) (string, error) {
		return worksheet.GetFormula(cellRange)
	})
}

// CreateHTMLTable creates a table data in HTML format
func createHTMLTable(startCol int, startRow int, endCol int, endRow int, extractor func(cellRange string) (string, error)) (*string, error) {
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
			table += fmt.Sprintf("<td>%s</td>", strings.ReplaceAll(html.EscapeString(value), "\n", "<br>"))
		}
		table += "</tr>\n"
	}

	table += "</table>"
	return &table, nil
}
