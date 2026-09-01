package excel

import (
	"errors"
	"fmt"
	"testing"

	"github.com/xuri/excelize/v2"
)

// newSharedGroupFile builds a workbook whose sheet holds one shared formula
// group, the shape excelize writes for a fill-down: the master carries the
// text and a ref, the members only carry the si.
func newSharedGroupFile(t *testing.T, sheet string, master string, ref string, formula string) *excelize.File {
	t.Helper()
	file := excelize.NewFile()
	if sheet != "Sheet1" {
		if _, err := file.NewSheet(sheet); err != nil {
			t.Fatal(err)
		}
		if err := file.DeleteSheet("Sheet1"); err != nil {
			t.Fatal(err)
		}
	}
	for row := 1; row <= 5; row++ {
		if err := file.SetCellInt(sheet, fmt.Sprintf("A%d", row), int64(row)); err != nil {
			t.Fatal(err)
		}
	}
	shared := excelize.STCellFormulaTypeShared
	if err := file.SetCellFormula(sheet, master, formula, excelize.FormulaOpts{Type: &shared, Ref: &ref}); err != nil {
		t.Fatal(err)
	}
	return file
}

// TestSharedFormulaMemberWriteIsIgnored pins the excelize behaviour that makes
// eachSheetCell necessary. If this ever starts failing, excelize has learned
// to split a shared group on a member write and rewriteFormulas can stop
// caring about groups.
func TestSharedFormulaMemberWriteIsIgnored(t *testing.T) {
	file := newSharedGroupFile(t, "Sheet1", "B1", "B1:B5", "A1*2")
	defer file.Close()

	if err := file.SetCellFormula("Sheet1", "B3", "A3*99"); err != nil {
		t.Fatal(err)
	}
	got, err := file.GetCellFormula("Sheet1", "B3")
	if err != nil {
		t.Fatal(err)
	}
	if got != "A3*2" {
		t.Fatalf("expected the member write to be swallowed by the group, got %q", got)
	}
}

// TestSharedFormulaIsInvisibleToAccessors pins the other half: the exported
// accessors report nothing that separates a group member from a plain cell.
func TestSharedFormulaIsInvisibleToAccessors(t *testing.T) {
	file := newSharedGroupFile(t, "Sheet1", "B1", "B1:B5", "A1*2")
	defer file.Close()

	master, err := file.GetCellFormula("Sheet1", "B1")
	if err != nil {
		t.Fatal(err)
	}
	member, err := file.GetCellFormula("Sheet1", "B4")
	if err != nil {
		t.Fatal(err)
	}
	if master != "A1*2" || member != "A4*2" {
		t.Fatalf("expected shifted formulas, got master %q and member %q", master, member)
	}
	cellType, err := file.GetCellType("Sheet1", "B4")
	if err != nil {
		t.Fatal(err)
	}
	if cellType != excelize.CellTypeUnset {
		t.Fatalf("expected a member to report no cell type, got %v", cellType)
	}
}

func TestEachSheetCellReportsSharedGroupStructure(t *testing.T) {
	file := newSharedGroupFile(t, "Sheet1", "B1", "B1:B5", "A1*2")
	defer file.Close()

	workbook := &ExcelizeExcel{file: file}
	masters, members := 0, 0
	err := workbook.eachSheetCell("Sheet1", func(cell sheetCell) error {
		if cell.Formula == nil || cell.Formula.Type != excelize.STCellFormulaTypeShared {
			return nil
		}
		if cell.Formula.Si == nil {
			t.Fatalf("shared cell %s has no si", cell.Ref)
		}
		if cell.Formula.Text != "" {
			masters++
		} else {
			members++
		}
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}
	if masters != 1 || members != 4 {
		t.Fatalf("expected 1 master and 4 members, got %d and %d", masters, members)
	}
}

func TestEachSheetCellStopsOnVisitError(t *testing.T) {
	file := newSharedGroupFile(t, "Sheet1", "B1", "B1:B5", "A1*2")
	defer file.Close()

	sentinel := errors.New("stop")
	visited := 0
	err := (&ExcelizeExcel{file: file}).eachSheetCell("Sheet1", func(sheetCell) error {
		visited++
		return sentinel
	})
	if !errors.Is(err, sentinel) {
		t.Fatalf("expected the visit error to surface, got %v", err)
	}
	if visited != 1 {
		t.Fatalf("expected the walk to stop at the first cell, visited %d", visited)
	}
}

func TestEachSheetCellRejectsUnknownSheet(t *testing.T) {
	file := excelize.NewFile()
	defer file.Close()

	err := (&ExcelizeExcel{file: file}).eachSheetCell("Nope", func(sheetCell) error {
		t.Fatal("no cell should be visited")
		return nil
	})
	if err == nil {
		t.Fatal("expected an error for a sheet that does not exist")
	}
}

func TestEachSheetCellReportsChartSheets(t *testing.T) {
	file := excelize.NewFile()
	defer file.Close()

	if err := file.SetCellFormula("Sheet1", "B1", "SUM(A1:A5)"); err != nil {
		t.Fatal(err)
	}
	chartType := excelize.Col
	if err := file.AddChartSheet("Chart", &excelize.Chart{
		Type:   chartType,
		Series: []excelize.ChartSeries{{Values: "Sheet1!$A$1:$A$5"}},
	}); err != nil {
		t.Fatal(err)
	}

	// Without the check the chart sheet reaches excelize's cell accessors,
	// which reject it, and the whole rewrite fails.
	if _, err := file.GetSheetDimension("Chart"); err == nil {
		t.Fatal("expected excelize to reject a chart sheet")
	}

	err := (&ExcelizeExcel{file: file}).eachSheetCell("Chart", func(sheetCell) error {
		t.Fatal("a chart sheet has no cells to visit")
		return nil
	})
	if !errors.Is(err, errNotWorksheet) {
		t.Fatalf("expected errNotWorksheet for a chart sheet, got %v", err)
	}
}

// TestRewriteFormulasSkipsChartSheets covers the whole path: GetSheetList hands
// out chart sheets too, and a workbook holding one must still rename.
func TestRewriteFormulasSkipsChartSheets(t *testing.T) {
	file := excelize.NewFile()
	defer file.Close()

	if _, err := file.NewSheet("Source"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellInt("Source", "A1", 7); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellFormula("Sheet1", "B1", "Source!A1*2"); err != nil {
		t.Fatal(err)
	}
	if err := file.AddChartSheet("Chart", &excelize.Chart{
		Type:   excelize.Col,
		Series: []excelize.ChartSeries{{Values: "Sheet1!$A$1:$A$5"}},
	}); err != nil {
		t.Fatal(err)
	}

	workbook := &ExcelizeExcel{file: file}
	if _, err := workbook.RenameSheet("Source", "Renamed"); err != nil {
		t.Fatalf("rename failed on a workbook with a chart sheet: %v", err)
	}
	got, err := file.GetCellFormula("Sheet1", "B1")
	if err != nil {
		t.Fatal(err)
	}
	if got != "Renamed!A1*2" {
		t.Fatalf("expected the formula to follow the rename, got %q", got)
	}
}

// TestRenameKeepsSharedGroupIntact guards the read path: a rewrite that changes
// nothing in a group must leave the group exactly as it was.
func TestRenameKeepsSharedGroupIntact(t *testing.T) {
	file := newSharedGroupFile(t, "Data", "B1", "B1:B5", "A1*2")
	defer file.Close()

	if _, err := file.NewSheet("Source"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellFormula("Data", "C1", "SUM(Source!A1:A2)"); err != nil {
		t.Fatal(err)
	}

	workbook := &ExcelizeExcel{file: file}
	if _, err := workbook.RenameSheet("Source", "Renamed"); err != nil {
		t.Fatal(err)
	}

	for row := 1; row <= 5; row++ {
		cell := fmt.Sprintf("B%d", row)
		got, err := file.GetCellFormula("Data", cell)
		if err != nil {
			t.Fatal(err)
		}
		want := fmt.Sprintf("A%d*2", row)
		if got != want {
			t.Fatalf("shared group damaged at %s: got %q, want %q", cell, got, want)
		}
	}
}

// TestRenameRewritesSharedGroup covers the write path: every member follows the
// rewritten master, so the group survives with only the master changed.
func TestRenameRewritesSharedGroup(t *testing.T) {
	file := newSharedGroupFile(t, "Data", "B1", "B1:B5", "Source!A1*2")
	defer file.Close()

	if _, err := file.NewSheet("Source"); err != nil {
		t.Fatal(err)
	}

	workbook := &ExcelizeExcel{file: file}
	if _, err := workbook.RenameSheet("Source", "Renamed"); err != nil {
		t.Fatal(err)
	}

	for row := 1; row <= 5; row++ {
		cell := fmt.Sprintf("B%d", row)
		got, err := file.GetCellFormula("Data", cell)
		if err != nil {
			t.Fatal(err)
		}
		if got == "" {
			t.Fatalf("formula at %s was wiped", cell)
		}
		if got == fmt.Sprintf("Source!A%d*2", row) {
			t.Fatalf("formula at %s still refers to the old sheet name: %q", cell, got)
		}
	}
}

// TestDeleteRowsSplitsSharedGroup covers the diverging case: one member falls
// inside the deleted rows, so the group has to be written out cell by cell
// instead of losing every member with the master.
func TestDeleteRowsSplitsSharedGroup(t *testing.T) {
	file := newSharedGroupFile(t, "Sheet1", "B1", "B1:B5", "A1*2")
	defer file.Close()

	sheet := &ExcelizeWorksheet{file: file, sheetName: "Sheet1"}
	if err := sheet.DeleteRows(1, 1); err != nil {
		t.Fatal(err)
	}

	// Rows 2 to 5 moved up one, and their formulas moved with them.
	for row := 1; row <= 4; row++ {
		cell := fmt.Sprintf("B%d", row)
		got, err := file.GetCellFormula("Sheet1", cell)
		if err != nil {
			t.Fatal(err)
		}
		if got == "" {
			t.Fatalf("formula at %s was wiped by deleting the master's row", cell)
		}
	}
}
