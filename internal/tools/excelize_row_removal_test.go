package tools

import (
	"fmt"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

// TestExcelizeRemoveRowBehaviour pins the raw excelize behaviour that
// ExcelizeWorksheet.DeleteRows works around. If an excelize upgrade changes
// any of this, the workaround needs a second look, so fail loudly here rather
// than in a tool test that only sees the combined result.
func TestExcelizeRemoveRowBehaviour(t *testing.T) {
	path := filepath.Join(t.TempDir(), "raw.xlsx")
	file := excelize.NewFile()
	sheet := file.GetSheetName(0)
	for row := 1; row <= 20; row++ {
		if err := file.SetCellInt(sheet, fmt.Sprintf("A%d", row), int64(row)); err != nil {
			t.Fatal(err)
		}
	}
	// All formula cells sit in rows 1-4, above the block that gets deleted,
	// so the cells themselves survive.
	formulas := map[string]string{
		"C1": "=A20",
		"C2": "=A6",
		"C3": "=SUM(A1:A20)",
		"D1": "=SUM(A6:A10)",
		"D2": "=A2",
		"D3": "=A1+A6+A20",
		"D4": "=SUM(A16:A20)",
	}
	for cell, formula := range formulas {
		if err := file.SetCellFormula(sheet, cell, formula); err != nil {
			t.Fatal(err)
		}
	}
	if _, err := file.NewSheet("Other"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellFormula("Other", "A1", "=Sheet1!A6"); err != nil {
		t.Fatal(err)
	}
	formulaType, ref := excelize.STCellFormulaTypeShared, "E1:E3"
	if err := file.SetCellFormula(sheet, "E1", "A1*2", excelize.FormulaOpts{Type: &formulaType, Ref: &ref}); err != nil {
		t.Fatal(err)
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	file.Close()

	reopened, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()
	// Rows 5-15.
	for i := 0; i < 11; i++ {
		if err := reopened.RemoveRow(sheet, 5); err != nil {
			t.Fatal(err)
		}
	}

	// GetCellFormula strips the leading "=", and a reference into the deleted
	// block is shifted onto a surviving row instead of becoming #REF!.
	expected := map[string]string{
		"C1": "A9",
		"C2": "A4",
		"C3": "SUM(A1:A9)",
		"D1": "SUM(A4:A4)",
		"D2": "A2",
		"D3": "A1+A4+A9",
		"D4": "SUM(A5:A9)",
	}
	for cell, want := range expected {
		if got, _ := reopened.GetCellFormula(sheet, cell); got != want {
			t.Errorf("%s: %s became %q, excelize used to produce %q", cell, formulas[cell], got, want)
		}
	}
	if got, _ := reopened.GetCellFormula("Other", "A1"); got != "Sheet1!A4" {
		t.Errorf("cross-sheet reference became %q, excelize used to produce Sheet1!A4", got)
	}

	// Removing the master cell of a shared formula group wipes the formula of
	// every member, not just the master.
	if err := reopened.RemoveRow(sheet, 1); err != nil {
		t.Fatal(err)
	}
	for _, cell := range []string{"E1", "E2"} {
		if got, _ := reopened.GetCellFormula(sheet, cell); got != "" {
			t.Errorf("%s = %q, excelize used to wipe the whole shared group with its master", cell, got)
		}
	}
}
