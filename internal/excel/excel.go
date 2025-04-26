package excel

import (
	"github.com/xuri/excelize/v2"
)

type Excel interface {
	// GetSheetNames returns the names of all sheets in the Excel file.
	GetSheetNames() ([]string, error)
	// FindSheet finds a sheet by its name and returns a Worksheet.
	FindSheet(sheetName string) (Worksheet, error)
	// Save saves the Excel file.
	Save() error
}

type Worksheet interface {
	// Name returns the name of the worksheet.
	Name() (string, error)
	// SetValue sets a value in the specified cell.
	SetValue(cell string, value any) error
	// SetFormula sets a formula in the specified cell.
	SetFormula(cell string, formula string) error
	// GetValue gets the value from the specified cell.
	GetValue(cell string) (string, error)
	// GetFormula gets the formula from the specified cell.
	GetFormula(cell string) (string, error)
	// GetDimention gets the dimension of the worksheet.
	GetDimention() (string, error)
	// GetPagingStrategy returns the paging strategy for the worksheet.
	// The pageSize parameter is used to determine the max size of each page.
	GetPagingStrategy(pageSize int) (PagingStrategy, error)
	// CapturePicture returns base64 encoded image data of the specified range.
	CapturePicture(captureRange string) (string, error)
}

// OpenFile opens an Excel file and returns an Excel interface.
// It first tries to open the file using OLE automation, and if that fails,
// it tries to using the excelize library.
func OpenFile(absoluteFilePath string) (Excel, func(), error) {
	ole, releaseFn, err := NewExcelOle(absoluteFilePath)
	if err == nil {
		return ole, releaseFn, nil
	}
	// If OLE fails, try Excelize
	workbook, err := excelize.OpenFile(absoluteFilePath)
	if err != nil {
		return nil, nil, err
	}
	excelize := NewExcelizeExcel(workbook)
	return excelize, func() {
		workbook.Close()
	}, nil
}
