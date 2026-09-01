package excel

import (
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"path"
	"path/filepath"
	"sort"
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

func (e *ExcelizeExcel) SheetNames() ([]string, error) {
	return e.file.GetSheetList(), nil
}

func (e *ExcelizeExcel) RenameSheet(oldSheetName string, newSheetName string) ([]string, error) {
	oldIndex, err := e.file.GetSheetIndex(oldSheetName)
	if err != nil || oldIndex < 0 {
		return nil, fmt.Errorf("sheet not found: %s", oldSheetName)
	}
	if oldSheetName == newSheetName {
		return nil, nil
	}
	// GetSheetIndex matches case-insensitively, so an index that is not the
	// source sheet means a different sheet already owns the target name.
	if existingIndex, err := e.file.GetSheetIndex(newSheetName); err == nil && existingIndex >= 0 && existingIndex != oldIndex {
		return nil, fmt.Errorf("a sheet named [%s] already exists", newSheetName)
	}

	// Formulas go first: excelize does not validate references, so pointing
	// them at a name that does not exist yet is fine, and the writes then
	// never have to care which name the sheet currently carries.
	err = e.rewriteFormulas(func(sheet string, cell string, formula string) (string, bool) {
		return ReplaceSheetNameInFormula(formula, oldSheetName, newSheetName)
	})
	if err != nil {
		return nil, err
	}
	// SetSheetName preserves all cell content, styles, merges, column widths
	// and row heights; it only edits the sheet's entry in workbook.xml and the
	// defined names that point at it.
	return nil, e.file.SetSheetName(oldSheetName, newSheetName)
}

// sheetCell mirrors the parts of a worksheet <c> element that matter for
// formula rewriting and used-range detection.
type sheetCell struct {
	Ref     string    `xml:"r,attr"`
	Value   string    `xml:"v"`
	Inline  *struct{} `xml:"is"`
	Formula *struct {
		Type string `xml:"t,attr"`
		Si   *int   `xml:"si,attr"`
		Text string `xml:",chardata"`
	} `xml:"f"`
}

func (c sheetCell) hasContent() bool {
	return c.Value != "" || c.Inline != nil || c.Formula != nil
}

// errNotWorksheet marks a sheet that holds no cells, such as a chart sheet or
// a dialog sheet. GetSheetList returns those alongside real worksheets.
var errNotWorksheet = errors.New("sheet is not a worksheet")

// eachSheetCell visits every cell of the sheet as excelize currently holds it
// in memory, in row order.
//
// The reason for reading the worksheet part rather than using the accessors is
// that no exported excelize API reports the shared formula attributes t, si
// and ref. GetCellFormula resolves a group member to the master's formula
// shifted to the member's own position and drops every trace of the group;
// GetCellType reports a member as CellTypeUnset, and Rows only yields values.
// The distinction matters because writing into a member leaves the group in
// place and Excel keeps reading the master, so the write is silently lost. It
// is not about speed: GetCellFormula binary searches the rows and is cheap.
//
// Cells are handed over one at a time instead of collected, because a
// worksheet can hold hundreds of thousands of them while both callers only
// keep a handful.
func (e *ExcelizeExcel) eachSheetCell(sheet string, visit func(sheetCell) error) error {
	path, err := e.sheetXMLPath(sheet)
	if err != nil {
		return err
	}
	// Any accessor makes excelize parse the sheet into memory.
	if _, err := e.file.GetSheetDimension(sheet); err != nil {
		return err
	}
	worksheet, ok := e.file.Sheet.Load(path)
	if !ok {
		return fmt.Errorf("worksheet part %s of sheet [%s] is not loaded", path, sheet)
	}
	// The parsed worksheet is an unexported excelize type, so serialising it
	// is the only way to look inside.
	raw, err := xml.Marshal(worksheet)
	if err != nil {
		return err
	}
	decoder := xml.NewDecoder(bytes.NewReader(raw))
	inSheetData := false
	for {
		token, err := decoder.Token()
		if errors.Is(err, io.EOF) {
			return nil
		}
		if err != nil {
			return err
		}
		switch element := token.(type) {
		case xml.StartElement:
			switch {
			case element.Name.Local == "sheetData":
				inSheetData = true
			case inSheetData && element.Name.Local == "c":
				var cell sheetCell
				if err := decoder.DecodeElement(&cell, &element); err != nil {
					return err
				}
				if err := visit(cell); err != nil {
					return err
				}
			}
		case xml.EndElement:
			// Everything after the cells is of no interest, and a worksheet
			// carries a lot of it.
			if element.Name.Local == "sheetData" {
				return nil
			}
		}
	}
}

// sheetXMLPath resolves a sheet name to its worksheet part in the package by
// way of workbook.xml and the workbook relationships, as excelize does
// privately. Sheets that are not worksheets report errNotWorksheet.
func (e *ExcelizeExcel) sheetXMLPath(sheet string) (string, error) {
	relationshipID := ""
	for _, entry := range e.file.WorkBook.Sheets.Sheet {
		if strings.EqualFold(entry.Name, sheet) {
			relationshipID = entry.ID
		}
	}
	relationships, ok := e.file.Relationships.Load("xl/_rels/workbook.xml.rels")
	if relationshipID == "" || !ok {
		return "", fmt.Errorf("cannot locate the worksheet part of sheet [%s]", sheet)
	}
	raw, err := xml.Marshal(relationships)
	if err != nil {
		return "", err
	}
	var parsed struct {
		Relationships []struct {
			ID     string `xml:"Id,attr"`
			Type   string `xml:"Type,attr"`
			Target string `xml:"Target,attr"`
		} `xml:"Relationship"`
	}
	if err := xml.Unmarshal(raw, &parsed); err != nil {
		return "", err
	}
	for _, relationship := range parsed.Relationships {
		if relationship.ID != relationshipID {
			continue
		}
		// The relationship type is what separates a worksheet from a chart or
		// dialog sheet; those have no cells to read.
		if !strings.HasSuffix(relationship.Type, "/worksheet") {
			return "", fmt.Errorf("%w: [%s]", errNotWorksheet, sheet)
		}
		// Targets are relative to xl/ unless they start at the package root.
		if strings.HasPrefix(relationship.Target, "/") {
			return strings.TrimPrefix(relationship.Target, "/"), nil
		}
		return path.Join("xl", relationship.Target), nil
	}
	return "", fmt.Errorf("cannot locate the worksheet part of sheet [%s]", sheet)
}

// rewriteFormulas offers every formula in the workbook to rewrite and stores
// the ones it changes, an empty result meaning the formula is removed. A
// rewrite that never reports a change makes this a read-only scan.
//
// Shared formula groups are expanded to their members, since excelize stores
// only the master's text and the rest inherit it by offset. When every member
// still follows the rewritten master the group survives with only the master
// changed; otherwise it is written back as individual formulas. Writing into
// a member directly is never right: the text would sit next to the inherited
// formula, and Excel then reads the master's formula in every member,
// unshifted.
func (e *ExcelizeExcel) rewriteFormulas(rewrite func(sheet string, cell string, formula string) (string, bool)) error {
	for _, sheet := range e.file.GetSheetList() {
		apply := func(cell string, formula string) (string, bool) {
			updated, changed := rewrite(sheet, cell, formula)
			if !changed {
				return formula, false
			}
			return updated, updated != formula
		}
		masters := map[int]sheetCell{}
		groups := map[int][]string{}
		var groupOrder []int
		// The cells are read from a snapshot of the worksheet, so writing back
		// while visiting them is safe.
		err := e.eachSheetCell(sheet, func(cell sheetCell) error {
			if cell.Formula == nil || cell.Formula.Type == excelize.STCellFormulaTypeDataTable {
				return nil
			}
			if cell.Formula.Type == excelize.STCellFormulaTypeShared && cell.Formula.Si != nil {
				si := *cell.Formula.Si
				if cell.Formula.Text != "" {
					masters[si] = cell
				}
				if _, seen := groups[si]; !seen {
					groupOrder = append(groupOrder, si)
				}
				groups[si] = append(groups[si], cell.Ref)
				return nil
			}
			if updated, changed := apply(cell.Ref, cell.Formula.Text); changed {
				if err := e.file.SetCellFormula(sheet, cell.Ref, updated); err != nil {
					return fmt.Errorf("failed to update formula in %s!%s: %w", sheet, cell.Ref, err)
				}
			}
			return nil
		})
		if errors.Is(err, errNotWorksheet) {
			continue
		}
		if err != nil {
			return err
		}

		for _, si := range groupOrder {
			master, ok := masters[si]
			if !ok {
				// A group without a master has no formula text anywhere.
				continue
			}
			masterCol, masterRow, err := excelize.CellNameToCoordinates(master.Ref)
			if err != nil {
				return err
			}
			masterUpdated, _ := apply(master.Ref, master.Formula.Text)
			members := groups[si]
			formulas := make([]string, len(members))
			anyChanged, uniform := false, true
			for i, member := range members {
				col, row, err := excelize.CellNameToCoordinates(member)
				if err != nil {
					return err
				}
				offsetCol, offsetRow := col-masterCol, row-masterRow
				updated, changed := apply(member, ShiftFormulaReferences(master.Formula.Text, offsetCol, offsetRow))
				formulas[i] = updated
				anyChanged = anyChanged || changed
				uniform = uniform && updated == ShiftFormulaReferences(masterUpdated, offsetCol, offsetRow)
			}
			if !anyChanged {
				continue
			}
			if uniform {
				// SetCellFormula on the master keeps the group's attributes.
				if err := e.file.SetCellFormula(sheet, master.Ref, masterUpdated); err != nil {
					return fmt.Errorf("failed to update formula in %s!%s: %w", sheet, master.Ref, err)
				}
				continue
			}
			// Clearing the master drops the inherited formula of every member,
			// after which each gets its own.
			if err := e.file.SetCellFormula(sheet, master.Ref, ""); err != nil {
				return err
			}
			for i, member := range members {
				if err := e.file.SetCellFormula(sheet, member, formulas[i]); err != nil {
					return fmt.Errorf("failed to update formula in %s!%s: %w", sheet, member, err)
				}
			}
		}
	}
	return nil
}

func (e *ExcelizeExcel) DeleteSheet(sheetName string, force bool) ([]string, error) {
	index, err := e.file.GetSheetIndex(sheetName)
	if err != nil || index < 0 {
		return nil, fmt.Errorf("sheet not found: %s", sheetName)
	}
	sheetList := e.file.GetSheetList()
	if len(sheetList) <= 1 {
		return nil, fmt.Errorf("cannot delete sheet [%s]: a workbook must keep at least one sheet", sheetName)
	}
	// Resolve to the stored casing so later comparisons are exact.
	sheetName = sheetList[index]

	// Formulas on the surviving sheets would turn into #REF! errors. Excelize
	// does not detect this, so refuse rather than hand back a broken workbook.
	var referencing []string
	err = e.rewriteFormulas(func(sheet string, cell string, formula string) (string, bool) {
		if sheet != sheetName && FormulaReferencesSheet(formula, sheetName) {
			referencing = append(referencing, sheet+"!"+cell)
		}
		return "", false
	})
	if err != nil {
		return nil, err
	}
	if len(referencing) > 0 && !force {
		shown := referencing
		if len(shown) > 10 {
			shown = shown[:10]
		}
		return nil, fmt.Errorf(
			"refusing to delete sheet [%s]: %d formula cell(s) reference it and would become #REF! errors (%s). Pass force=true to delete anyway",
			sheetName, len(referencing), strings.Join(shown, ", "))
	}

	var warnings []string
	if len(referencing) > 0 {
		warnings = append(warnings, fmt.Sprintf(
			"%d formula cell(s) referenced [%s] and are now broken: %s",
			len(referencing), sheetName, strings.Join(referencing, ", ")))
	}

	// Excelize drops defined names scoped to the deleted sheet, but leaves
	// workbook-scoped ones pointing at it. Remove those too.
	for _, definedName := range e.file.GetDefinedName() {
		if definedName.Scope != "Workbook" {
			continue
		}
		if !FormulaReferencesSheet(definedName.RefersTo, sheetName) {
			continue
		}
		if err := e.file.DeleteDefinedName(&excelize.DefinedName{
			Name:  definedName.Name,
			Scope: definedName.Scope,
		}); err != nil {
			return nil, fmt.Errorf("failed to remove defined name [%s] referring to sheet [%s]: %w", definedName.Name, sheetName, err)
		}
		warnings = append(warnings, fmt.Sprintf("removed defined name [%s] which referred to the deleted sheet", definedName.Name))
	}

	if err := e.file.DeleteSheet(sheetName); err != nil {
		return nil, err
	}
	return warnings, nil
}

func (e *ExcelizeExcel) MoveSheet(sheetName string, index int) error {
	sheetList := e.file.GetSheetList()
	currentIndex, err := e.file.GetSheetIndex(sheetName)
	if err != nil || currentIndex < 0 {
		return fmt.Errorf("sheet not found: %s", sheetName)
	}
	if index < 0 || index >= len(sheetList) {
		return fmt.Errorf("index %d is out of range: the workbook has %d sheets (valid indexes are 0-%d)", index, len(sheetList), len(sheetList)-1)
	}
	if index == currentIndex {
		return nil
	}

	// Build the wanted order, then realise it with excelize's "move before"
	// primitive. Walking backwards keeps the already-placed suffix intact.
	sheetName = sheetList[currentIndex]
	desired := make([]string, 0, len(sheetList))
	for i, name := range sheetList {
		if i != currentIndex {
			desired = append(desired, name)
		}
	}
	desired = append(desired[:index], append([]string{sheetName}, desired[index:]...)...)

	for i := len(desired) - 2; i >= 0; i-- {
		if err := e.file.MoveSheet(desired[i], desired[i+1]); err != nil {
			return err
		}
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
	// Force a one-time full recalculation on next file open. Excelize does not
	// update cached <v> values for cells whose formulas depend on writes we
	// just made; in workbooks with calcMode="manual" this leaves stale values
	// visible in Excel/LibreOffice/Numbers until the user manually recalculates.
	fullCalc := true
	if err := w.file.SetCalcProps(&excelize.CalcPropsOptions{FullCalcOnLoad: &fullCalc}); err != nil {
		return err
	}
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
	// Capture existing style ID before writing, so that we can restore
	// number formats / fonts / fills / borders that excelize would otherwise
	// strip when the new value's type differs from the previous cell's type
	// (e.g. writing a numeric/formula result into a cell that previously held
	// a string, or vice versa). GetCellStyle returns 0 (default) when the cell
	// has no explicit style, in which case the restore is a no-op.
	styleID, styleErr := w.file.GetCellStyle(w.sheetName, cell)
	if err := w.file.SetCellValue(w.sheetName, cell, value); err != nil {
		return err
	}
	if styleErr == nil && styleID != 0 {
		if err := w.file.SetCellStyle(w.sheetName, cell, cell, styleID); err != nil {
			return fmt.Errorf("failed to restore cell style: %w", err)
		}
	}
	if err := w.updateDimension(cell); err != nil {
		return fmt.Errorf("failed to update dimension: %w", err)
	}
	return nil
}

func (w *ExcelizeWorksheet) SetFormula(cell string, formula string) error {
	// Capture and restore style ID across the formula write — see SetValue for
	// rationale. SetCellFormula is especially prone to stripping numFmt when
	// the previous cell type differs.
	styleID, styleErr := w.file.GetCellStyle(w.sheetName, cell)
	if err := w.file.SetCellFormula(w.sheetName, cell, formula); err != nil {
		return err
	}
	if styleErr == nil && styleID != 0 {
		if err := w.file.SetCellStyle(w.sheetName, cell, cell, styleID); err != nil {
			return fmt.Errorf("failed to restore cell style: %w", err)
		}
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

func (w *ExcelizeWorksheet) GetRawValue(cell string) (string, error) {
	return w.file.GetCellValue(w.sheetName, cell, excelize.Options{RawCellValue: true})
}

func (w *ExcelizeWorksheet) GetCellType(cell string) (string, error) {
	cellType, err := w.file.GetCellType(w.sheetName, cell)
	if err != nil {
		return "", err
	}
	switch cellType {
	case excelize.CellTypeBool:
		return "bool", nil
	case excelize.CellTypeDate:
		return "date", nil
	case excelize.CellTypeNumber:
		return "number", nil
	case excelize.CellTypeError:
		return "error", nil
	case excelize.CellTypeFormula:
		return "formula", nil
	case excelize.CellTypeSharedString, excelize.CellTypeInlineString:
		return "string", nil
	default:
		return "unknown", nil
	}
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

func (w *ExcelizeWorksheet) SetCellStyle(cell string, style *CellStyle) error {
	excelizeStyle := convertCellStyleToExcelizeStyle(style)

	styleID, err := w.file.NewStyle(excelizeStyle)
	if err != nil {
		return fmt.Errorf("failed to create style: %w", err)
	}

	if err := w.file.SetCellStyle(w.sheetName, cell, cell, styleID); err != nil {
		return fmt.Errorf("failed to set cell style: %w", err)
	}

	return nil
}

func (w *ExcelizeWorksheet) GetMergedCells() ([]string, error) {
	merged, err := w.file.GetMergeCells(w.sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get merged cells: %w", err)
	}
	ranges := make([]string, 0, len(merged))
	for _, m := range merged {
		ranges = append(ranges, m.GetStartAxis()+":"+m.GetEndAxis())
	}
	return ranges, nil
}

func (w *ExcelizeWorksheet) GetColumnWidths(startCol int, endCol int) (map[string]float64, error) {
	widths := make(map[string]float64)
	for col := startCol; col <= endCol; col++ {
		name, err := excelize.ColumnNumberToName(col)
		if err != nil {
			return nil, err
		}
		width, err := w.file.GetColWidth(w.sheetName, name)
		if err != nil {
			return nil, fmt.Errorf("failed to get width of column %s: %w", name, err)
		}
		widths[name] = width
	}
	return widths, nil
}

func (w *ExcelizeWorksheet) DeleteRows(startRow int, endRow int) error {
	if startRow < 1 {
		return fmt.Errorf("startRow must be 1 or greater, got %d", startRow)
	}
	if endRow < startRow {
		return fmt.Errorf("endRow (%d) must not be smaller than startRow (%d)", endRow, startRow)
	}
	// Excelize shifts a reference to a deleted row onto whichever row lands in
	// its place, which is silently wrong; Excel breaks it to #REF!. Do that
	// first, while the rows still exist. Formulas inside the deleted rows are
	// dropped here rather than left to RemoveRow, because excelize wipes every
	// member of a shared group when the master cell goes.
	workbook := &ExcelizeExcel{file: w.file}
	err := workbook.rewriteFormulas(func(sheet string, cell string, formula string) (string, bool) {
		if strings.EqualFold(sheet, w.sheetName) {
			if _, row, err := excelize.CellNameToCoordinates(cell); err == nil && row >= startRow && row <= endRow {
				return "", true
			}
		}
		return BreakReferencesToRows(formula, sheet, w.sheetName, startRow, endRow)
	})
	if err != nil {
		return err
	}
	// Excelize removes one row at a time and adjusts merges, conditional
	// formats, data validations, defined names and formulas on each call, so
	// the row to remove stays startRow as the rows below shift up.
	for i := 0; i < endRow-startRow+1; i++ {
		if err := w.file.RemoveRow(w.sheetName, startRow); err != nil {
			return err
		}
	}
	return w.recalculateDimension()
}

func (w *ExcelizeWorksheet) InsertRows(beforeRow int, count int) error {
	if beforeRow < 1 {
		return fmt.Errorf("beforeRow must be 1 or greater, got %d", beforeRow)
	}
	if count < 1 {
		return fmt.Errorf("count must be 1 or greater, got %d", count)
	}
	if err := w.file.InsertRows(w.sheetName, beforeRow, count); err != nil {
		return err
	}
	return w.recalculateDimension()
}

// recalculateDimension recomputes the used range from the sheet's contents.
// updateDimension only ever grows the range, but deleting rows has to shrink
// it too, otherwise paging keeps walking over rows that no longer exist.
func (w *ExcelizeWorksheet) recalculateDimension() error {
	minCol, minRow, maxCol, maxRow := 0, 0, 0, 0
	err := (&ExcelizeExcel{file: w.file}).eachSheetCell(w.sheetName, func(cell sheetCell) error {
		if !cell.hasContent() {
			return nil
		}
		column, row, err := excelize.CellNameToCoordinates(cell.Ref)
		if err != nil {
			return err
		}
		if minCol == 0 || column < minCol {
			minCol = column
		}
		if column > maxCol {
			maxCol = column
		}
		if minRow == 0 || row < minRow {
			minRow = row
		}
		if row > maxRow {
			maxRow = row
		}
		return nil
	})
	if errors.Is(err, errNotWorksheet) {
		return nil
	}
	if err != nil {
		return err
	}
	if maxRow == 0 {
		return w.file.SetSheetDimension(w.sheetName, "A1")
	}
	start, err := excelize.CoordinatesToCellName(minCol, minRow)
	if err != nil {
		return err
	}
	end, err := excelize.CoordinatesToCellName(maxCol, maxRow)
	if err != nil {
		return err
	}
	return w.file.SetSheetDimension(w.sheetName, start+":"+end)
}

func (w *ExcelizeWorksheet) GetConditionalFormatRanges() ([]string, error) {
	formats, err := w.file.GetConditionalFormats(w.sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get conditional formats: %w", err)
	}
	ranges := make([]string, 0, len(formats))
	for reference := range formats {
		ranges = append(ranges, reference)
	}
	sort.Strings(ranges)
	return ranges, nil
}

func (w *ExcelizeWorksheet) GetDataValidationRanges() ([]string, error) {
	validations, err := w.file.GetDataValidations(w.sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get data validations: %w", err)
	}
	ranges := make([]string, 0, len(validations))
	for _, validation := range validations {
		if validation != nil && validation.Sqref != "" {
			ranges = append(ranges, validation.Sqref)
		}
	}
	sort.Strings(ranges)
	return ranges, nil
}

func convertCellStyleToExcelizeStyle(style *CellStyle) *excelize.Style {
	result := &excelize.Style{}

	// Border
	if len(style.Border) > 0 {
		borders := make([]excelize.Border, len(style.Border))
		for i, border := range style.Border {
			excelizeBorder := excelize.Border{
				Type: border.Type.String(),
			}
			if border.Color != "" {
				excelizeBorder.Color = strings.TrimPrefix(border.Color, "#")
			}
			excelizeBorder.Style = borderStyleNameToInt(border.Style)
			borders[i] = excelizeBorder
		}
		result.Border = borders
	}

	// Font
	if style.Font != nil {
		font := &excelize.Font{}
		if style.Font.Bold != nil {
			font.Bold = *style.Font.Bold
		}
		if style.Font.Italic != nil {
			font.Italic = *style.Font.Italic
		}
		if style.Font.Underline != nil {
			font.Underline = style.Font.Underline.String()
		}
		if style.Font.Size != nil && *style.Font.Size > 0 {
			font.Size = float64(*style.Font.Size)
		}
		if style.Font.Strike != nil {
			font.Strike = *style.Font.Strike
		}
		if style.Font.Color != nil && *style.Font.Color != "" {
			font.Color = strings.TrimPrefix(*style.Font.Color, "#")
		}
		if style.Font.VertAlign != nil {
			font.VertAlign = style.Font.VertAlign.String()
		}
		result.Font = font
	}

	// Fill
	if style.Fill != nil {
		fill := excelize.Fill{}
		if style.Fill.Type != "" {
			fill.Type = style.Fill.Type.String()
		}
		fill.Pattern = fillPatternNameToInt(style.Fill.Pattern)
		if len(style.Fill.Color) > 0 {
			colors := make([]string, len(style.Fill.Color))
			for i, color := range style.Fill.Color {
				colors[i] = strings.TrimPrefix(color, "#")
			}
			fill.Color = colors
		}
		if style.Fill.Shading != nil {
			fill.Shading = fillShadingNameToInt(*style.Fill.Shading)
		}
		result.Fill = fill
	}

	// Alignment
	if style.Alignment != nil {
		alignment := &excelize.Alignment{}
		if style.Alignment.Horizontal != nil {
			alignment.Horizontal = *style.Alignment.Horizontal
		}
		if style.Alignment.Vertical != nil {
			alignment.Vertical = *style.Alignment.Vertical
		}
		if style.Alignment.WrapText != nil {
			alignment.WrapText = *style.Alignment.WrapText
		}
		if style.Alignment.ShrinkToFit != nil {
			alignment.ShrinkToFit = *style.Alignment.ShrinkToFit
		}
		if style.Alignment.TextRotation != nil {
			alignment.TextRotation = *style.Alignment.TextRotation
		}
		if style.Alignment.Indent != nil {
			alignment.Indent = *style.Alignment.Indent
		}
		result.Alignment = alignment
	}

	// NumFmt
	if style.NumFmt != nil && *style.NumFmt != "" {
		result.CustomNumFmt = style.NumFmt
	}

	// DecimalPlaces
	if style.DecimalPlaces != nil && *style.DecimalPlaces > 0 {
		result.DecimalPlaces = style.DecimalPlaces
	}

	return result
}

func convertExcelizeStyleToCellStyle(style *excelize.Style) *CellStyle {
	result := &CellStyle{}

	// Border
	if len(style.Border) > 0 {
		var borders []Border
		for _, border := range style.Border {
			borderStyle := Border{
				Type: BorderType(border.Type),
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
			font.Bold = &style.Font.Bold
		}
		if style.Font.Italic {
			font.Italic = &style.Font.Italic
		}
		if style.Font.Underline != "" {
			underline := FontUnderline(style.Font.Underline)
			font.Underline = &underline
		}
		if style.Font.Size > 0 {
			size := int(style.Font.Size)
			font.Size = &size
		}
		if style.Font.Strike {
			font.Strike = &style.Font.Strike
		}
		if style.Font.Color != "" {
			color := "#" + strings.ToUpper(style.Font.Color)
			font.Color = &color
		}
		if style.Font.VertAlign != "" {
			vertAlign := FontVertAlign(style.Font.VertAlign)
			font.VertAlign = &vertAlign
		}
		if font.Bold != nil || font.Italic != nil || font.Underline != nil || font.Size != nil || font.Strike != nil || font.Color != nil || font.VertAlign != nil {
			result.Font = font
		}
	}

	// Fill
	if style.Fill.Type != "" || style.Fill.Pattern != 0 || len(style.Fill.Color) > 0 {
		fill := &FillStyle{}
		if style.Fill.Type != "" {
			fill.Type = FillType(style.Fill.Type)
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
			shading := intToFillShadingName(style.Fill.Shading)
			fill.Shading = &shading
		}
		if fill.Type != "" || fill.Pattern != FillPatternNone || len(fill.Color) > 0 || fill.Shading != nil {
			result.Fill = fill
		}
	}

	// Alignment
	if style.Alignment != nil {
		alignment := &AlignmentStyle{}
		hasAlignment := false
		if style.Alignment.Horizontal != "" {
			alignment.Horizontal = &style.Alignment.Horizontal
			hasAlignment = true
		}
		if style.Alignment.Vertical != "" {
			alignment.Vertical = &style.Alignment.Vertical
			hasAlignment = true
		}
		if style.Alignment.WrapText {
			alignment.WrapText = &style.Alignment.WrapText
			hasAlignment = true
		}
		if style.Alignment.ShrinkToFit {
			alignment.ShrinkToFit = &style.Alignment.ShrinkToFit
			hasAlignment = true
		}
		if style.Alignment.TextRotation != 0 {
			alignment.TextRotation = &style.Alignment.TextRotation
			hasAlignment = true
		}
		if style.Alignment.Indent != 0 {
			alignment.Indent = &style.Alignment.Indent
			hasAlignment = true
		}
		if hasAlignment {
			result.Alignment = alignment
		}
	}

	// NumFmt
	if style.CustomNumFmt != nil && *style.CustomNumFmt != "" {
		result.NumFmt = style.CustomNumFmt
	}

	// DecimalPlaces
	if style.DecimalPlaces != nil && *style.DecimalPlaces != 0 {
		result.DecimalPlaces = style.DecimalPlaces
	}

	return result
}

func intToBorderStyleName(style int) BorderStyle {
	styles := map[int]BorderStyle{
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

func intToFillPatternName(pattern int) FillPattern {
	patterns := map[int]FillPattern{
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

func intToFillShadingName(shading int) FillShading {
	shadings := map[int]FillShading{
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

func borderStyleNameToInt(style BorderStyle) int {
	styles := map[BorderStyle]int{
		BorderStyleNone:             0,
		BorderStyleContinuous:       1,
		BorderStyleDash:             3,
		BorderStyleDot:              4,
		BorderStyleDouble:           6,
		BorderStyleDashDot:          8,
		BorderStyleDashDotDot:       9,
		BorderStyleSlantDashDot:     10,
		BorderStyleMediumDashDot:    12,
		BorderStyleMediumDashDotDot: 13,
	}
	if value, exists := styles[style]; exists {
		return value
	}
	return 1
}

func fillPatternNameToInt(pattern FillPattern) int {
	patterns := map[FillPattern]int{
		FillPatternNone:            0,
		FillPatternSolid:           1,
		FillPatternMediumGray:      2,
		FillPatternDarkGray:        3,
		FillPatternLightGray:       4,
		FillPatternDarkHorizontal:  5,
		FillPatternDarkVertical:    6,
		FillPatternDarkDown:        7,
		FillPatternDarkUp:          8,
		FillPatternDarkGrid:        9,
		FillPatternDarkTrellis:     10,
		FillPatternLightHorizontal: 11,
		FillPatternLightVertical:   12,
		FillPatternLightDown:       13,
		FillPatternLightUp:         14,
		FillPatternLightGrid:       15,
		FillPatternLightTrellis:    16,
		FillPatternGray125:         17,
		FillPatternGray0625:        18,
	}
	if value, exists := patterns[pattern]; exists {
		return value
	}
	return 0
}

func fillShadingNameToInt(shading FillShading) int {
	shadings := map[FillShading]int{
		FillShadingHorizontal:   0,
		FillShadingVertical:     1,
		FillShadingDiagonalDown: 2,
		FillShadingDiagonalUp:   3,
		FillShadingFromCenter:   4,
		FillShadingFromCorner:   5,
	}
	if value, exists := shadings[shading]; exists {
		return value
	}
	return 0
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
