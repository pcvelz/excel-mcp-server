package excel

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

type ExcelizeExcel struct {
	file *excelize.File
}

func NewExcelizeExcel(file *excelize.File) Excel {
	return &ExcelizeExcel{file: file}
}

func (e *ExcelizeExcel) GetBackendName() string {
	return "excelize"
}

func (e *ExcelizeExcel) FindSheet(sheetName string) (Worksheet, error) {
	index, err := e.file.GetSheetIndex(sheetName)
	if err != nil {
		return nil, fmt.Errorf("sheet not found: %s", sheetName)
	}
	if index < 0 {
		return nil, fmt.Errorf("sheet not found: %s", sheetName)
	}
	return &ExcelizeWorksheet{file: e.file, sheetName: sheetName}, nil
}

func (e *ExcelizeExcel) CreateNewSheet(sheetName string) error {
	_, err := e.file.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to create new sheet: %w", err)
	}
	return nil
}

func (e *ExcelizeExcel) CopySheet(srcSheetName string, destSheetName string) error {
	srcIndex, err := e.file.GetSheetIndex(srcSheetName)
	if srcIndex < 0 {
		return fmt.Errorf("source sheet not found: %s", srcSheetName)
	}
	if err != nil {
		return err
	}
	destIndex, err := e.file.NewSheet(destSheetName)
	if err != nil {
		return fmt.Errorf("failed to create destination sheet: %w", err)
	}
	if err := e.file.CopySheet(srcIndex, destIndex); err != nil {
		return fmt.Errorf("failed to copy sheet: %w", err)
	}
	srcNext := e.file.GetSheetList()[srcIndex+1]
	if srcNext != srcSheetName {
		e.file.MoveSheet(destSheetName, srcNext)
	}
	return nil
}

func (e *ExcelizeExcel) GetSheets() ([]Worksheet, error) {
	sheetList := e.file.GetSheetList()
	worksheets := make([]Worksheet, len(sheetList))
	for i, sheetName := range sheetList {
		worksheets[i] = &ExcelizeWorksheet{file: e.file, sheetName: sheetName}
	}
	return worksheets, nil
}

// SaveExcelize saves the Excel file to the specified path.
// Excelize's Save method restricts the file path length to 207 characters,
// but since this limitation has been relaxed in some environments,
// we ignore this restriction.
// https://github.com/qax-os/excelize/blob/v2.9.0/file.go#L71-L73
func (w *ExcelizeExcel) Save() error {
	file, err := os.OpenFile(filepath.Clean(w.file.Path), os.O_WRONLY|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if err != nil {
		return err
	}
	defer file.Close()
	return w.file.Write(file)
}

type ExcelizeWorksheet struct {
	file      *excelize.File
	sheetName string
}

func (w *ExcelizeWorksheet) Release() {
	// No resources to release in excelize
}

func (w *ExcelizeWorksheet) Name() (string, error) {
	return w.sheetName, nil
}

func (w *ExcelizeWorksheet) GetTables() ([]Table, error) {
	tables, err := w.file.GetTables(w.sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get tables: %w", err)
	}
	tableList := make([]Table, len(tables))
	for i, table := range tables {
		tableList[i] = Table{
			Name:  table.Name,
			Range: NormalizeRange(table.Range),
		}
	}
	return tableList, nil
}

func (w *ExcelizeWorksheet) GetPivotTables() ([]PivotTable, error) {
	pivotTables, err := w.file.GetPivotTables(w.sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get pivot tables: %w", err)
	}
	pivotTableList := make([]PivotTable, len(pivotTables))
	for i, pivotTable := range pivotTables {
		pivotTableList[i] = PivotTable{
			Name:  pivotTable.Name,
			Range: NormalizeRange(pivotTable.PivotTableRange),
		}
	}
	return pivotTableList, nil
}

func (w *ExcelizeWorksheet) SetValue(cell string, value any) error {
	if err := w.file.SetCellValue(w.sheetName, cell, value); err != nil {
		return err
	}
	if err := w.updateDimension(cell); err != nil {
		return fmt.Errorf("failed to update dimension: %w", err)
	}
	return nil
}

func (w *ExcelizeWorksheet) SetFormula(cell string, formula string) error {
	if err := w.file.SetCellFormula(w.sheetName, cell, formula); err != nil {
		return err
	}
	if err := w.updateDimension(cell); err != nil {
		return fmt.Errorf("failed to update dimension: %w", err)
	}
	return nil
}

func (w *ExcelizeWorksheet) GetValue(cell string) (string, error) {
	value, err := w.file.GetCellValue(w.sheetName, cell)
	if err != nil {
		return "", err
	}
	if value == "" {
		// try to get calculated value
		formula, err := w.file.GetCellFormula(w.sheetName, cell)
		if err != nil {
			return "", fmt.Errorf("failed to get formula: %w", err)
		}
		if formula != "" {
			return w.file.CalcCellValue(w.sheetName, cell)
		}
	}
	return value, nil
}

func (w *ExcelizeWorksheet) GetFormula(cell string) (string, error) {
	formula, err := w.file.GetCellFormula(w.sheetName, cell)
	if err != nil {
		return "", fmt.Errorf("failed to get formula: %w", err)
	}
	if formula == "" {
		// fallback
		return w.GetValue(cell)
	}
	if !strings.HasPrefix(formula, "=") {
		formula = "=" + formula
	}
	return formula, nil
}

func (w *ExcelizeWorksheet) GetDimention() (string, error) {
	return w.file.GetSheetDimension(w.sheetName)
}

func (w *ExcelizeWorksheet) GetPagingStrategy(pageSize int) (PagingStrategy, error) {
	return NewExcelizeFixedSizePagingStrategy(pageSize, w)
}

func (w *ExcelizeWorksheet) CapturePicture(captureRange string) (string, error) {
	return "", fmt.Errorf("CapturePicture is not supported in Excelize")
}

func (w *ExcelizeWorksheet) AddTable(tableRange, tableName string) error {
	enable := true
	if err := w.file.AddTable(w.sheetName, &excelize.Table{
		Range:             tableRange,
		Name:              tableName,
		StyleName:         "TableStyleMedium2",
		ShowColumnStripes: true,
		ShowFirstColumn:   false,
		ShowHeaderRow:     &enable,
		ShowLastColumn:    false,
		ShowRowStripes:    &enable,
	}); err != nil {
		return err
	}
	return nil
}

func (w *ExcelizeWorksheet) GetCellStyle(cell string) (*CellStyle, error) {
	styleID, err := w.file.GetCellStyle(w.sheetName, cell)
	if err != nil {
		return nil, fmt.Errorf("failed to get cell style: %w", err)
	}

	style, err := w.file.GetStyle(styleID)
	if err != nil {
		return nil, fmt.Errorf("failed to get style details: %w", err)
	}

	return convertExcelizeStyleToCellStyle(style), nil
}

func convertExcelizeStyleToCellStyle(style *excelize.Style) *CellStyle {
	result := &CellStyle{}

	// Border
	if len(style.Border) > 0 {
		var borders []BorderStyle
		for _, border := range style.Border {
			borderStyle := BorderStyle{
				Type: border.Type,
			}
			if border.Color != "" {
				borderStyle.Color = "#" + strings.ToUpper(border.Color)
			}
			if border.Style != 0 {
				borderStyle.Style = intToBorderStyleName(border.Style)
			}
			borders = append(borders, borderStyle)
		}
		if len(borders) > 0 {
			result.Border = borders
		}
	}

	// Font
	if style.Font != nil {
		font := &FontStyle{}
		if style.Font.Bold {
			font.Bold = true
		}
		if style.Font.Italic {
			font.Italic = true
		}
		if style.Font.Underline != "" {
			font.Underline = style.Font.Underline
		}
		if style.Font.Size > 0 {
			font.Size = int(style.Font.Size)
		}
		if style.Font.Strike {
			font.Strike = true
		}
		if style.Font.Color != "" {
			font.Color = "#" + strings.ToUpper(style.Font.Color)
		}
		if style.Font.VertAlign != "" {
			font.VertAlign = style.Font.VertAlign
		}
		if font.Bold || font.Italic || font.Underline != "" || font.Size > 0 || font.Strike || font.Color != "" || font.VertAlign != "" {
			result.Font = font
		}
	}

	// Fill
	if style.Fill.Type != "" || style.Fill.Pattern != 0 || len(style.Fill.Color) > 0 {
		fill := &FillStyle{}
		if style.Fill.Type != "" {
			fill.Type = style.Fill.Type
		}
		if style.Fill.Pattern != 0 {
			fill.Pattern = intToFillPatternName(style.Fill.Pattern)
		}
		if len(style.Fill.Color) > 0 {
			var colors []string
			for _, color := range style.Fill.Color {
				if color != "" {
					colors = append(colors, "#"+strings.ToUpper(color))
				}
			}
			if len(colors) > 0 {
				fill.Color = colors
			}
		}
		if style.Fill.Shading != 0 {
			fill.Shading = intToFillShadingName(style.Fill.Shading)
		}
		if fill.Type != "" || fill.Pattern != FillPatternNone || len(fill.Color) > 0 || fill.Shading != FillShadingHorizontal {
			result.Fill = fill
		}
	}

	// NumFmt
	if style.CustomNumFmt != nil && *style.CustomNumFmt != "" {
		result.NumFmt = *style.CustomNumFmt
	}

	// DecimalPlaces
	if style.DecimalPlaces != nil && *style.DecimalPlaces != 0 {
		result.DecimalPlaces = *style.DecimalPlaces
	}
	
	return result
}

func intToBorderStyleName(style int) BorderStyleName {
	styles := map[int]BorderStyleName{
		0:  BorderStyleNone,
		1:  BorderStyleContinuous,
		2:  BorderStyleContinuous,
		3:  BorderStyleDash,
		4:  BorderStyleDot,
		5:  BorderStyleContinuous,
		6:  BorderStyleDouble,
		7:  BorderStyleContinuous,
		8:  BorderStyleDashDot,
		9:  BorderStyleDashDotDot,
		10: BorderStyleSlantDashDot,
		11: BorderStyleContinuous,
		12: BorderStyleMediumDashDot,
		13: BorderStyleMediumDashDotDot,
	}
	if name, exists := styles[style]; exists {
		return name
	}
	return BorderStyleContinuous
}

func intToFillPatternName(pattern int) FillPatternName {
	patterns := map[int]FillPatternName{
		0:  FillPatternNone,
		1:  FillPatternSolid,
		2:  FillPatternMediumGray,
		3:  FillPatternDarkGray,
		4:  FillPatternLightGray,
		5:  FillPatternDarkHorizontal,
		6:  FillPatternDarkVertical,
		7:  FillPatternDarkDown,
		8:  FillPatternDarkUp,
		9:  FillPatternDarkGrid,
		10: FillPatternDarkTrellis,
		11: FillPatternLightHorizontal,
		12: FillPatternLightVertical,
		13: FillPatternLightDown,
		14: FillPatternLightUp,
		15: FillPatternLightGrid,
		16: FillPatternLightTrellis,
		17: FillPatternGray125,
		18: FillPatternGray0625,
	}
	if name, exists := patterns[pattern]; exists {
		return name
	}
	return FillPatternNone
}

func intToFillShadingName(shading int) FillShadingName {
	shadings := map[int]FillShadingName{
		0: FillShadingHorizontal,
		1: FillShadingVertical,
		2: FillShadingDiagonalDown,
		3: FillShadingDiagonalUp,
		4: FillShadingFromCenter,
		5: FillShadingFromCorner,
	}
	if name, exists := shadings[shading]; exists {
		return name
	}
	return FillShadingHorizontal
}

// updateDimention updates the dimension of the worksheet after a cell is updated.
func (w *ExcelizeWorksheet) updateDimension(updatedCell string) error {
	dimension, err := w.file.GetSheetDimension(w.sheetName)
	if err != nil {
		return err
	}
	startCol, startRow, endCol, endRow, err := ParseRange(dimension)
	if err != nil {
		return err
	}
	updatedCol, updatedRow, err := excelize.CellNameToCoordinates(updatedCell)
	if err != nil {
		return err
	}
	if startCol > updatedCol {
		startCol = updatedCol
	}
	if endCol < updatedCol {
		endCol = updatedCol
	}
	if startRow > updatedRow {
		startRow = updatedRow
	}
	if endRow < updatedRow {
		endRow = updatedRow
	}
	startRange, err := excelize.CoordinatesToCellName(startCol, startRow)
	if err != nil {
		return err
	}
	endRange, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		return err
	}
	updatedDimension := fmt.Sprintf("%s:%s", startRange, endRange)
	return w.file.SetSheetDimension(w.sheetName, updatedDimension)
}
