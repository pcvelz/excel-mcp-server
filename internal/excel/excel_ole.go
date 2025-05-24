package excel

import (
	"bufio"
	"bytes"
	"fmt"
	"io"
	"path/filepath"
	"strings"

	"encoding/base64"

	"github.com/skanehira/clipboard-image"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type OleExcel struct {
	workbook *ole.IDispatch
}

type OleWorksheet struct {
	worksheet *ole.IDispatch
}

func NewExcelOle(absolutePath string) (*OleExcel, func(), error) {
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

func NewExcelOleWithNewObject(absolutePath string) (*OleExcel, func(), error) {
	ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)

	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		return nil, func() {}, err
	}
	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, func() {}, err
	}
	workbooks := oleutil.MustGetProperty(excel, "Workbooks").ToIDispatch()
	workbook, err := oleutil.CallMethod(workbooks, "Open", absolutePath)
	if err != nil {
		return nil, func() {}, err
	}
	w := workbook.ToIDispatch()
	return &OleExcel{workbook: w}, func() {
		w.Release()
		workbooks.Release()
		excel.Release()
		oleutil.CallMethod(excel, "Close")
		ole.CoUninitialize()
	}, nil
}

func (o *OleExcel) GetBackendName() string {
	return "ole"
}

func (o *OleExcel) GetSheets() ([]Worksheet, error) {
	worksheets := oleutil.MustGetProperty(o.workbook, "Worksheets").ToIDispatch()
	defer worksheets.Release()

	count := int(oleutil.MustGetProperty(worksheets, "Count").Val)
	worksheetList := make([]Worksheet, count)

	for i := 1; i <= count; i++ {
		worksheet := oleutil.MustGetProperty(worksheets, "Item", i).ToIDispatch()
		worksheetList[i-1] = &OleWorksheet{
			worksheet: worksheet,
		}
	}
	return worksheetList, nil
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
	}

	return nil, fmt.Errorf("sheet not found: %s", sheetName)
}

func (o *OleExcel) CreateNewSheet(sheetName string) error {
	activeWorksheet := oleutil.MustGetProperty(o.workbook, "ActiveSheet").ToIDispatch()
	defer activeWorksheet.Release()
	activeWorksheetIndex := oleutil.MustGetProperty(activeWorksheet, "Index").Val
	worksheets := oleutil.MustGetProperty(o.workbook, "Worksheets").ToIDispatch()
	defer worksheets.Release()

	_, err := oleutil.CallMethod(worksheets, "Add", nil, activeWorksheet)
	if err != nil {
		return err
	}

	worksheet := oleutil.MustGetProperty(worksheets, "Item", activeWorksheetIndex+1).ToIDispatch()
	defer worksheet.Release()

	_, err = oleutil.PutProperty(worksheet, "Name", sheetName)
	if err != nil {
		return err
	}

	return nil
}

func (o *OleExcel) CopySheet(srcSheetName string, dstSheetName string) error {
	worksheets := oleutil.MustGetProperty(o.workbook, "Worksheets").ToIDispatch()
	defer worksheets.Release()

	srcSheetVariant, err := oleutil.GetProperty(worksheets, "Item", srcSheetName)
	if err != nil {
		return fmt.Errorf("faild to get sheet: %w", err)
	}
	srcSheet := srcSheetVariant.ToIDispatch()
	defer srcSheet.Release()
	srcSheetIndex := oleutil.MustGetProperty(srcSheet, "Index").Val

	_, err = oleutil.CallMethod(srcSheet, "Copy", nil, srcSheet)
	if err != nil {
		return err
	}

	dstSheetVariant, err := oleutil.GetProperty(worksheets, "Item", srcSheetIndex+1)
	if err != nil {
		return fmt.Errorf("failed to get copied sheet: %w", err)
	}
	dstSheet := dstSheetVariant.ToIDispatch()
	defer dstSheet.Release()

	_, err = oleutil.PutProperty(dstSheet, "Name", dstSheetName)
	if err != nil {
		return err
	}

	return nil
}

func (o *OleExcel) Save() error {
	_, err := oleutil.CallMethod(o.workbook, "Save")
	if err != nil {
		return err
	}
	return nil
}

func (o *OleWorksheet) Release() {
	o.worksheet.Release()
}

func (o *OleWorksheet) Name() (string, error) {
	name := oleutil.MustGetProperty(o.worksheet, "Name").ToString()
	return name, nil
}

func (o *OleWorksheet) GetTables() ([]Table, error) {
	tables := oleutil.MustGetProperty(o.worksheet, "ListObjects").ToIDispatch()
	defer tables.Release()
	count := int(oleutil.MustGetProperty(tables, "Count").Val)
	tableList := make([]Table, count)
	for i := 1; i <= count; i++ {
		table := oleutil.MustGetProperty(tables, "Item", i).ToIDispatch()
		defer table.Release()
		name := oleutil.MustGetProperty(table, "Name").ToString()
		defer table.Release()
		tableRange := oleutil.MustGetProperty(table, "Range").ToIDispatch()
		defer tableRange.Release()
		tableList[i-1] = Table{
			Name:  name,
			Range: NormalizeRange(oleutil.MustGetProperty(tableRange, "Address").ToString()),
		}
	}
	return tableList, nil
}

func (o *OleWorksheet) GetPivotTables() ([]PivotTable, error) {
	pivotTables := oleutil.MustGetProperty(o.worksheet, "PivotTables").ToIDispatch()
	defer pivotTables.Release()
	count := int(oleutil.MustGetProperty(pivotTables, "Count").Val)
	pivotTableList := make([]PivotTable, count)
	for i := 1; i <= count; i++ {
		pivotTable := oleutil.MustGetProperty(pivotTables, "Item", i).ToIDispatch()
		defer pivotTable.Release()
		name := oleutil.MustGetProperty(pivotTable, "Name").ToString()
		pivotTableRange := oleutil.MustGetProperty(pivotTable, "TableRange1").ToIDispatch()
		defer pivotTableRange.Release()
		pivotTableList[i-1] = PivotTable{
			Name:  name,
			Range: NormalizeRange(oleutil.MustGetProperty(pivotTableRange, "Address").ToString()),
		}
	}
	return pivotTableList, nil
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
	value := oleutil.MustGetProperty(range_, "Text").Value()
	switch v := value.(type) {
	case string:
		return v, nil
	case nil:
		return "", nil
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
	return NormalizeRange(dimension), nil
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

func (o *OleWorksheet) CapturePicture(captureRange string) (string, error) {
	r := oleutil.MustGetProperty(o.worksheet, "Range", captureRange).ToIDispatch()
	defer r.Release()
	_, err := oleutil.CallMethod(
		r,
		"CopyPicture",
		int(1), // xlScreen (https://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel.xlpictureappearance?view=excel-pia)
		int(2), // xlBitmap (https://learn.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.excel.xlcopypictureformat?view=excel-pia)
	)
	if err != nil {
		return "", err
	}
	// Read the image from the clipboard
	buf := new(bytes.Buffer)
	bufWriter := bufio.NewWriter(buf)
	clipboardReader, err := clipboard.ReadFromClipboard()
	if err != nil {
		return "", fmt.Errorf("failed to read from clipboard: %w", err)
	}
	if _, err := io.Copy(bufWriter, clipboardReader); err != nil {
		return "", fmt.Errorf("failed to copy clipboard data: %w", err)
	}
	if err := bufWriter.Flush(); err != nil {
		return "", fmt.Errorf("failed to flush buffer: %w", err)
	}
	return base64.StdEncoding.EncodeToString(buf.Bytes()), nil
}

func (o *OleWorksheet) AddTable(tableRange string, tableName string) error {
	tables := oleutil.MustGetProperty(o.worksheet, "ListObjects").ToIDispatch()
	defer tables.Release()

	// https://learn.microsoft.com/ja-jp/office/vba/api/excel.listobjects.add
	tableVar, err := oleutil.CallMethod(
		tables,
		"Add",
		int(1), // xlSrcRange (https://learn.microsoft.com/ja-jp/office/vba/api/excel.xllistobjectsourcetype)
		tableRange,
		nil,
		int(0), // xlYes (https://learn.microsoft.com/ja-jp/office/vba/api/excel.xlyesnoguess)
	)
	if err != nil {
		return err
	}
	table := tableVar.ToIDispatch()
	defer table.Release()
	_, err = oleutil.PutProperty(table, "Name", tableName)
	if err != nil {
		return err
	}
	return err
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
