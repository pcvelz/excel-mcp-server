package excel

import (
	"fmt"
	"math"
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

func (e *ExcelizeExcel) GetSheetNames() ([]string, error) {
	sheetList := e.file.GetSheetList()
	return sheetList, nil
}

func (w *ExcelizeExcel) Save() error {
	file, err := os.OpenFile(filepath.Clean(w.file.Path), os.O_WRONLY|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if err != nil {
		return err
	}
	defer file.Close()
	if err := w.file.Write(file); err != nil {
		return err
	}
	if err := w.file.Save(); err != nil {
		return err
	}
	return nil
}

type ExcelizeWorksheet struct {
	file      *excelize.File
	sheetName string
}

func (w *ExcelizeWorksheet) Name() (string, error) {
	return w.sheetName, nil
}

func (w *ExcelizeWorksheet) SetValue(cell string, value any) error {
	return w.file.SetCellValue(w.sheetName, cell, value)
}

func (w *ExcelizeWorksheet) SetFormula(cell string, formula string) error {
	return w.file.SetCellFormula(w.sheetName, cell, formula)
}

func (w *ExcelizeWorksheet) GetValue(cell string) (string, error) {
	return w.file.GetCellValue(w.sheetName, cell)
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
	minRow, maxRow := math.MaxInt, 0
	minCol, maxCol := math.MaxInt, 0

	cols, err := w.file.Cols(w.sheetName)
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

func (w *ExcelizeWorksheet) GetPagingStrategy(pageSize int) (PagingStrategy, error) {
	return NewExcelizeFixedSizePagingStrategy(pageSize, w)
}

func (w *ExcelizeWorksheet) CapturePicture(captureRange string) (string, error) {
	return "", fmt.Errorf("CapturePicture is not supported in Excelize")
}
