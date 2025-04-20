package excel

import (
	"fmt"
	"path/filepath"
	"strings"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type OleExcel struct {
	workbook *ole.IDispatch
}

type OleWorksheet struct {
	worksheet *ole.IDispatch
}

func NewExcelOle(absolutePath string) (Excel, func(), error) {
	ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)

	unknown, err := oleutil.GetActiveObject("Excel.Application")
	if err != nil {
		return nil, func() {}, err
	}
	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, func() {}, err
	}
	workbooks := oleutil.MustGetProperty(excel, "Workbooks").ToIDispatch()
	c := oleutil.MustGetProperty(workbooks, "Count").Val
	for i := 1; i <= int(c); i++ {
		workbook := oleutil.MustGetProperty(workbooks, "Item", i).ToIDispatch()
		fullName := oleutil.MustGetProperty(workbook, "FullName").ToString()
		if normalizePath(fullName) == normalizePath(absolutePath) {
			return &OleExcel{workbook: workbook}, func() {
				workbook.Release()
				workbooks.Release()
				excel.Release()
				ole.CoUninitialize()
			}, nil
		}
	}
	return nil, func() {}, fmt.Errorf("workbook not found: %s", absolutePath)
}

func (o *OleExcel) FindSheet(sheetName string) (Worksheet, error) {
	worksheets := oleutil.MustGetProperty(o.workbook, "Worksheets").ToIDispatch()
	defer worksheets.Release()

	count := int(oleutil.MustGetProperty(worksheets, "Count").Val)

	for i := 1; i <= count; i++ {
		worksheet := oleutil.MustGetProperty(worksheets, "Item", i).ToIDispatch()
		name := oleutil.MustGetProperty(worksheet, "Name").ToString()

		if name == sheetName {
			return &OleWorksheet{
				worksheet: worksheet,
			}, nil
		}
		worksheet.Release()
	}

	return nil, fmt.Errorf("sheet not found: %s", sheetName)
}

func (o *OleExcel) Save() error {
	_, err := oleutil.CallMethod(o.workbook, "Save")
	if err != nil {
		return err
	}
	return nil
}

func (o *OleWorksheet) Name() (string, error) {
	name := oleutil.MustGetProperty(o.worksheet, "Name").ToString()
	return name, nil
}

func (o *OleWorksheet) SetValue(cell string, value any) error {
	range_ := oleutil.MustGetProperty(o.worksheet, "Range", cell).ToIDispatch()
	defer range_.Release()
	_, err := oleutil.PutProperty(range_, "Value", value)
	return err
}

func (o *OleWorksheet) SetFormula(cell string, formula string) error {
	range_ := oleutil.MustGetProperty(o.worksheet, "Range", cell).ToIDispatch()
	defer range_.Release()
	_, err := oleutil.PutProperty(range_, "Formula", formula)
	return err
}

func (o *OleWorksheet) GetValue(cell string) (string, error) {
	range_ := oleutil.MustGetProperty(o.worksheet, "Range", cell).ToIDispatch()
	defer range_.Release()
	value := oleutil.MustGetProperty(range_, "Value").Value()
	switch v := value.(type) {
	case int, int8, int16, int32, int64:
		return fmt.Sprintf("%d", v), nil
	case uint, uint8, uint16, uint32, uint64:
		return fmt.Sprintf("%d", v), nil
	case float32, float64:
		return fmt.Sprintf("%g", v), nil
	case complex64, complex128:
		return fmt.Sprintf("%g", v), nil
	case []byte:
		return string(v), nil
	case bool:
		return fmt.Sprintf("%t", v), nil
	case string:
		return v, nil
	case time.Time:
		return v.Format(time.RFC3339), nil
	case nil:
		return "", nil
	case *ole.VARIANT:
		return v.ToString(), nil
	default: // Handle other types as needed
		return "", fmt.Errorf("unsupported type: %T", v)
	}
}

func (o *OleWorksheet) GetFormula(cell string) (string, error) {
	range_ := oleutil.MustGetProperty(o.worksheet, "Range", cell).ToIDispatch()
	defer range_.Release()
	formula := oleutil.MustGetProperty(range_, "Formula").ToString()
	return formula, nil
}

func (o *OleWorksheet) GetDimention() (string, error) {
	range_ := oleutil.MustGetProperty(o.worksheet, "UsedRange").ToIDispatch()
	defer range_.Release()
	dimension := oleutil.MustGetProperty(range_, "Address").ToString()
	return dimension, nil
}

func (o *OleWorksheet) GetPagingStrategy(pageSize int) (PagingStrategy, error) {
	return NewOlePagingStrategy(1000, o)
}

func (o *OleWorksheet) PrintArea() (string, error) {
	v, err := oleutil.GetProperty(o.worksheet, "PageSetup")
	if err != nil {
		return "", err
	}
	pageSetup := v.ToIDispatch()
	defer pageSetup.Release()

	printArea := oleutil.MustGetProperty(pageSetup, "PrintArea").ToString()
	return printArea, nil
}

func (o *OleWorksheet) HPageBreaks() ([]int, error) {
	v, err := oleutil.GetProperty(o.worksheet, "HPageBreaks")
	if err != nil {
		return nil, err
	}
	hPageBreaks := v.ToIDispatch()
	defer hPageBreaks.Release()

	count := int(oleutil.MustGetProperty(hPageBreaks, "Count").Val)
	pageBreaks := make([]int, count)
	for i := 1; i <= count; i++ {
		pageBreak := oleutil.MustGetProperty(hPageBreaks, "Item", i).ToIDispatch()
		defer pageBreak.Release()
		location := oleutil.MustGetProperty(pageBreak, "Location").ToIDispatch()
		defer location.Release()
		row := oleutil.MustGetProperty(location, "Row").Val
		pageBreaks[i-1] = int(row)
	}
	return pageBreaks, nil
}

func normalizePath(path string) string {
	// Normalize the volume name to uppercase
	vol := filepath.VolumeName(path)
	if vol == "" {
		return path
	}
	rest := path[len(vol):]
	return filepath.Clean(strings.ToUpper(vol) + rest)
}
