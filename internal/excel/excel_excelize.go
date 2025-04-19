package excel

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

type ExcelizeExcel struct {
	file *excelize.File
}

func NewExcelizeExcel(file *excelize.File) Excel {
	return &ExcelizeExcel{file: file}
}

func (e *ExcelizeExcel) FindSheet(sheetName string) (Worksheet, error) {
	// シートの存在確認
	index, err := e.file.GetSheetIndex(sheetName)
	if err != nil {
		return nil, fmt.Errorf("sheet not found: %s", sheetName)
	}
	if index < 0 {
		return nil, fmt.Errorf("sheet not found: %s", sheetName)
	}
	return &ExcelizeWorksheet{file: e.file, sheetName: sheetName}, nil
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

func (w *ExcelizeWorksheet) GetValue(cell string) (any, error) {
	return w.file.GetCellValue(w.sheetName, cell)
}

func (w *ExcelizeWorksheet) GetFormula(cell string) (string, error) {
	return w.file.GetCellFormula(w.sheetName, cell)
}
