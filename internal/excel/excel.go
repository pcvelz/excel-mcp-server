package excel

import (
	"github.com/xuri/excelize/v2"
)

type Excel interface {
	GetSheetNames() ([]string, error)
	FindSheet(sheetName string) (Worksheet, error)
	Save() error
}

type Worksheet interface {
	Name() (string, error)
	SetValue(cell string, value any) error
	SetFormula(cell string, formula string) error
	GetValue(cell string) (string, error)
	GetFormula(cell string) (string, error)
	GetDimention() (string, error)
	GetPagingStrategy(pageSize int) (PagingStrategy, error)
	CapturePicture(captureRange string) (string, error)
}

func OpenFile(absoluteFilePath string) (Excel, func(), error) {
	ole, releaseFn, err := NewExcelOle(absoluteFilePath)
	if err == nil {
		return ole, releaseFn, nil
	}
	workbook, err := excelize.OpenFile(absoluteFilePath)
	if err != nil {
		return nil, nil, err
	}
	excelize := NewExcelizeExcel(workbook)
	return excelize, func() {
		workbook.Close()
	}, nil
}
