package tools

import (
	"crypto/md5"
	"fmt"
	"html"
	"path/filepath"
	"slices"
	"strconv"
	"strings"

	"github.com/goccy/go-yaml"
	"github.com/xuri/excelize/v2"

	"github.com/negokaz/excel-mcp-server/internal/excel"

	z "github.com/Oudwins/zog"
)

type StyleRegistry struct {
	styles   map[string]*excel.CellStyle // styleID -> CellStyle
	hashToID map[string]string           // styleHash -> styleID
	counter  int
}

func NewStyleRegistry() *StyleRegistry {
	return &StyleRegistry{
		styles:   make(map[string]*excel.CellStyle),
		hashToID: make(map[string]string),
		counter:  0,
	}
}

func (sr *StyleRegistry) RegisterStyle(cellStyle *excel.CellStyle) string {
	if cellStyle == nil || sr.isEmptyStyle(cellStyle) {
		return ""
	}

	styleHash := sr.calculateStyleHash(cellStyle)

	if existingID, exists := sr.hashToID[styleHash]; exists {
		return existingID
	}

	sr.counter++
	styleID := fmt.Sprintf("s%d", sr.counter)
	sr.styles[styleID] = cellStyle
	sr.hashToID[styleHash] = styleID

	return styleID
}

func (sr *StyleRegistry) isEmptyStyle(style *excel.CellStyle) bool {
	if len(style.Border) > 0 || style.Font != nil || style.NumFmt != "" || style.DecimalPlaces != 0 {
		return false
	}
	if style.Fill != nil && style.Fill.Type != "" {
		return false
	}
	return true
}

func (sr *StyleRegistry) calculateStyleHash(cellStyle *excel.CellStyle) string {
	yamlBytes, err := yaml.MarshalWithOptions(cellStyle, yaml.Flow(true), yaml.OmitEmpty())
	if err != nil {
		return ""
	}

	hash := md5.Sum(yamlBytes)
	return fmt.Sprintf("%x", hash)[:8]
}

func (sr *StyleRegistry) GenerateStyleDefinitions() string {
	if len(sr.styles) == 0 {
		return ""
	}

	var result strings.Builder
	result.WriteString("<h2>Style Definitions</h2>\n")
	result.WriteString("<div class=\"style-definitions\">\n")

	var styleIDs []string
	for styleID := range sr.styles {
		styleIDs = append(styleIDs, styleID)
	}
	slices.SortFunc(styleIDs, func(a, b string) int {
		// styleID must have number suffix
		ai, _ := strconv.Atoi(a[1:])
		bi, _ := strconv.Atoi(b[1:])
		return ai - bi
	})

	for _, styleID := range styleIDs {
		cellStyle := sr.styles[styleID]
		yamlStr := convertCellStyleToYAMLFlow(cellStyle)
		if yamlStr != "" {
			result.WriteString(fmt.Sprintf("<code class=\"style language-yaml\" id=\"%s\">%s</code>\n", styleID, html.EscapeString(yamlStr)))
		}
	}

	result.WriteString("</div>\n\n")
	return result.String()
}

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
	return createHTMLTableWithStyle(startCol, startRow, endCol, endRow, extractor, nil)
}

func CreateHTMLTableOfValuesWithStyle(worksheet excel.Worksheet, startCol int, startRow int, endCol int, endRow int) (*string, error) {
	return createHTMLTableWithStyle(startCol, startRow, endCol, endRow,
		func(cellRange string) (string, error) {
			return worksheet.GetValue(cellRange)
		},
		func(cellRange string) (*excel.CellStyle, error) {
			return worksheet.GetCellStyle(cellRange)
		})
}

func CreateHTMLTableOfFormulaWithStyle(worksheet excel.Worksheet, startCol int, startRow int, endCol int, endRow int) (*string, error) {
	return createHTMLTableWithStyle(startCol, startRow, endCol, endRow,
		func(cellRange string) (string, error) {
			return worksheet.GetFormula(cellRange)
		},
		func(cellRange string) (*excel.CellStyle, error) {
			return worksheet.GetCellStyle(cellRange)
		})
}

func createHTMLTableWithStyle(startCol int, startRow int, endCol int, endRow int, extractor func(cellRange string) (string, error), styleExtractor func(cellRange string) (*excel.CellStyle, error)) (*string, error) {
	registry := NewStyleRegistry()

	// データとスタイルを収集
	var result strings.Builder
	result.WriteString("<table>\n<tr><th></th>")

	// 列アドレスの出力
	for col := startCol; col <= endCol; col++ {
		name, _ := excelize.ColumnNumberToName(col)
		result.WriteString(fmt.Sprintf("<th>%s</th>", name))
	}
	result.WriteString("</tr>\n")

	// データの出力とスタイル登録
	for row := startRow; row <= endRow; row++ {
		result.WriteString("<tr>")
		result.WriteString(fmt.Sprintf("<th>%d</th>", row))

		for col := startCol; col <= endCol; col++ {
			axis, _ := excelize.CoordinatesToCellName(col, row)
			value, _ := extractor(axis)

			var tdTag string
			if styleExtractor != nil {
				cellStyle, err := styleExtractor(axis)
				if err == nil && cellStyle != nil {
					styleID := registry.RegisterStyle(cellStyle)
					if styleID != "" {
						tdTag = fmt.Sprintf("<td style-ref=\"%s\">", styleID)
					} else {
						tdTag = "<td>"
					}
				} else {
					tdTag = "<td>"
				}
			} else {
				tdTag = "<td>"
			}

			result.WriteString(fmt.Sprintf("%s%s</td>", tdTag, strings.ReplaceAll(html.EscapeString(value), "\n", "<br>")))
		}
		result.WriteString("</tr>\n")
	}

	result.WriteString("</table>")

	// スタイル定義とテーブルを結合
	var finalResult strings.Builder
	styleDefinitions := registry.GenerateStyleDefinitions()
	if styleDefinitions != "" {
		finalResult.WriteString(styleDefinitions)
	}

	finalResult.WriteString("<h2>Sheet Data</h2>\n")
	finalResult.WriteString(result.String())

	finalResultStr := finalResult.String()
	return &finalResultStr, nil
}


func AbsolutePathTest() z.Test[*string] {
	return z.Test[*string]{
		Func: func(path *string, ctx z.Ctx) {
			if !filepath.IsAbs(*path) {
				ctx.AddIssue(ctx.Issue().SetMessage(fmt.Sprintf("Path '%s' is not absolute", *path)))
			}
		},
	}
}

func convertCellStyleToYAMLFlow(cellStyle *excel.CellStyle) string {
	if cellStyle == nil {
		return ""
	}
	yamlBytes, err := yaml.MarshalWithOptions(cellStyle, yaml.Flow(true), yaml.OmitEmpty())
	if err != nil {
		return ""
	}
	yamlStr := strings.TrimSpace(strings.ReplaceAll(string(yamlBytes), "\"", ""))
	return yamlStr
}

func convertStyleMapToYAMLFlow(styleMap map[string]any) string {
	if len(styleMap) == 0 {
		return ""
	}
	yamlBytes, err := yaml.MarshalWithOptions(styleMap, yaml.Flow(true), yaml.OmitEmpty())
	if err != nil {
		return ""
	}
	yamlStr := strings.TrimSpace(strings.ReplaceAll(string(yamlBytes), "\"", ""))
	return yamlStr
}
