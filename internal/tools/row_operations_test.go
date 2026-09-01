package tools

import (
	"fmt"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// TestDeleteRowsAcceptance is the case the row tools exist for: strip a sheet
// down to the rows that matter while every bit of formatting survives.
func TestDeleteRowsAcceptance(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "excel-mcp-tabtest.xlsx")
	buildTabTestWorkbook(t, path)

	ok(renameSheet(path, "Blad3", "Archief klanten"))
	ok(deleteRows(path, "Archief klanten", 3, 5))

	file, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer file.Close()

	sheet := "Archief klanten"
	if got, _ := file.GetCellValue(sheet, "A2"); got != "Username" {
		t.Errorf("A2 = %q, want the header Username", got)
	}
	for i, want := range []string{"actiefklant1", "actiefklant2", "actiefklant3"} {
		cell := fmt.Sprintf("A%d", 3+i)
		if got, _ := file.GetCellValue(sheet, cell); got != want {
			t.Errorf("%s = %q, want %s", cell, got, want)
		}
	}
	if got, _ := file.GetSheetDimension(sheet); got != "A2:D5" {
		t.Errorf("used range = %q, want A2:D5", got)
	}

	// The surviving client rows must still be green, bold, size 11.
	for row := 3; row <= 5; row++ {
		for _, column := range []string{"A", "B", "C", "D"} {
			cell := fmt.Sprintf("%s%d", column, row)
			styleID, err := file.GetCellStyle(sheet, cell)
			if err != nil {
				t.Fatal(err)
			}
			style, err := file.GetStyle(styleID)
			if err != nil {
				t.Fatal(err)
			}
			if style.Font == nil {
				t.Errorf("%s lost its font", cell)
				continue
			}
			if !style.Font.Bold || style.Font.Size != 11 {
				t.Errorf("%s font = bold:%v size:%v, want bold size 11", cell, style.Font.Bold, style.Font.Size)
			}
			if !strings.Contains(strings.ToUpper(style.Font.Color), "00B050") {
				t.Errorf("%s font colour = %q, want 00B050", cell, style.Font.Color)
			}
		}
	}

	// Sheet-level formatting survives.
	merged, err := file.GetMergeCells(sheet)
	if err != nil {
		t.Fatal(err)
	}
	foundMerge := false
	for _, m := range merged {
		if m.GetStartAxis() == "B2" && m.GetEndAxis() == "C2" {
			foundMerge = true
		}
	}
	if !foundMerge {
		t.Errorf("B2:C2 is no longer merged, merges = %v", merged)
	}
	width, err := file.GetColWidth(sheet, "B")
	if err != nil {
		t.Fatal(err)
	}
	if width != 40 {
		t.Errorf("column B width = %v, want 40", width)
	}
	styleID, _ := file.GetCellStyle(sheet, "D2")
	style, err := file.GetStyle(styleID)
	if err != nil {
		t.Fatal(err)
	}
	if style.Alignment == nil || !style.Alignment.WrapText {
		t.Error("D2 lost its wrapText alignment")
	}

	// The vertical centring that sat on A4 belonged to a deleted row: it must
	// be gone, not reassigned to a row that shifted up into its place.
	for row := 2; row <= 5; row++ {
		cell := fmt.Sprintf("A%d", row)
		styleID, _ := file.GetCellStyle(sheet, cell)
		style, err := file.GetStyle(styleID)
		if err != nil {
			t.Fatal(err)
		}
		if style.Alignment != nil && style.Alignment.Vertical == "center" {
			t.Errorf("%s unexpectedly carries the vertical centring from the deleted row", cell)
		}
	}
}

// TestDeleteRowsShiftsFormulas checks that formulas end up the way Excel
// leaves them after deleting rows 5-15: references below the block move up,
// ranges spanning it shrink, and anything pointing only into it breaks to
// #REF! instead of silently landing on another row.
func TestDeleteRowsShiftsFormulas(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "formulas.xlsx")

	file := excelize.NewFile()
	sheet := file.GetSheetName(0)
	for row := 1; row <= 20; row++ {
		if err := file.SetCellInt(sheet, fmt.Sprintf("A%d", row), int64(row)); err != nil {
			t.Fatal(err)
		}
	}
	// All formula cells sit in rows 1-4, above the deleted block.
	same := map[string]struct{ formula, want string }{
		"C1": {"=A20", "A9"},
		"C2": {"=A6", "#REF!"},
		"C3": {"=SUM(A1:A20)", "SUM(A1:A9)"},
		"C4": {"=SUM(A6:A10)", "SUM(#REF!)"},
		"D1": {"=A2", "A2"},
		"D2": {"=A1+A6+A20", "A1+#REF!+A9"},
		"D3": {"=SUM(A16:A20)", "SUM(A5:A9)"},
		"D4": {"=$A$6+A$7", "#REF!+#REF!"},
		"E1": {`=IF(A1="A6",A6,"A7")`, `IF(A1="A6",#REF!,"A7")`},
		"E2": {"=SUM(A4:A6)", "SUM(A4:A4)"},
		"E3": {"=SUM(5:15)+SUM(A:A)", "SUM(#REF!)+SUM(A:A)"},
		"E4": {"=[1]Sheet1!A6+LOG10(A6)", "[1]Sheet1!A6+LOG10(#REF!)"},
	}
	for cell, c := range same {
		if err := file.SetCellFormula(sheet, cell, c.formula); err != nil {
			t.Fatal(err)
		}
	}
	if _, err := file.NewSheet("Other"); err != nil {
		t.Fatal(err)
	}
	cross := map[string]struct{ formula, want string }{
		"A1": {"=Sheet1!A20", "Sheet1!A9"},
		"A2": {"=Sheet1!A6", "#REF!"},
		"A3": {"=SUM('Sheet1'!A6:A10)", "SUM(#REF!)"},
		"A4": {"=A6", "A6"},
		// excelize cannot parse Sheet1!#REF! and would then leave A20 alone,
		// which is why the qualifier is dropped from the broken reference.
		"A5": {"=Sheet1!A6+Sheet1!A20", "#REF!+Sheet1!A9"},
	}
	for cell, c := range cross {
		if err := file.SetCellFormula("Other", cell, c.formula); err != nil {
			t.Fatal(err)
		}
	}
	// A shared group filled down from F1: F1=A1*2 ... F20=A20*2. Excelize
	// stores only the master's text, so the members must be materialised
	// before their rows disappear or their references break.
	formulaType, ref := excelize.STCellFormulaTypeShared, "F1:F20"
	if err := file.SetCellFormula(sheet, "F1", "A1*2", excelize.FormulaOpts{Type: &formulaType, Ref: &ref}); err != nil {
		t.Fatal(err)
	}
	// A shared group above the block whose members all point into it.
	formulaType, ref = excelize.STCellFormulaTypeShared, "G1:G3"
	if err := file.SetCellFormula(sheet, "G1", "A6+$A$1", excelize.FormulaOpts{Type: &formulaType, Ref: &ref}); err != nil {
		t.Fatal(err)
	}
	names := map[string]struct{ refersTo, want string }{
		"Staart": {"Sheet1!$A$16:$A$20", "Sheet1!$A$5:$A$9"},
		"Midden": {"Sheet1!$A$6:$A$10", "#REF!"},
	}
	for name, n := range names {
		if err := file.SetDefinedName(&excelize.DefinedName{Name: name, RefersTo: n.refersTo}); err != nil {
			t.Fatal(err)
		}
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	file.Close()

	ok(deleteRows(path, sheet, 5, 15))

	reopened, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()

	for _, definedName := range reopened.GetDefinedName() {
		if n, ok := names[definedName.Name]; ok && definedName.RefersTo != n.want {
			t.Errorf("defined name %s: %s became %q, want %q", definedName.Name, n.refersTo, definedName.RefersTo, n.want)
		}
	}
	for cell, c := range same {
		if got, _ := reopened.GetCellFormula(sheet, cell); got != c.want {
			t.Errorf("%s: %s became %q, want %q", cell, c.formula, got, c.want)
		}
	}
	for cell, c := range cross {
		if got, _ := reopened.GetCellFormula("Other", cell); got != c.want {
			t.Errorf("Other!%s: %s became %q, want %q", cell, c.formula, got, c.want)
		}
	}
	for row, want := range map[int]string{1: "A1*2", 4: "A4*2", 5: "A5*2", 9: "A9*2"} {
		if got, _ := reopened.GetCellFormula(sheet, fmt.Sprintf("F%d", row)); got != want {
			t.Errorf("F%d = %q, want %q (shared group member)", row, got, want)
		}
	}
	if got, _ := reopened.GetCellFormula(sheet, "F10"); got != "" {
		t.Errorf("F10 = %q, want no formula: the group ended at F20, which is now row 9", got)
	}
	for row, want := range map[int]string{1: "#REF!+$A$1", 2: "#REF!+$A$1", 3: "#REF!+$A$1"} {
		if got, _ := reopened.GetCellFormula(sheet, fmt.Sprintf("G%d", row)); got != want {
			t.Errorf("G%d = %q, want %q", row, got, want)
		}
	}
}

// TestDeleteRowsKeepsUnaffectedSharedGroup makes sure a shared group that the
// deletion does not touch stays shared rather than being expanded needlessly.
func TestDeleteRowsKeepsUnaffectedSharedGroup(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "shared.xlsx")

	file := excelize.NewFile()
	sheet := file.GetSheetName(0)
	for row := 1; row <= 10; row++ {
		if err := file.SetCellInt(sheet, fmt.Sprintf("A%d", row), int64(row)); err != nil {
			t.Fatal(err)
		}
	}
	formulaType, ref := excelize.STCellFormulaTypeShared, "B1:B3"
	if err := file.SetCellFormula(sheet, "B1", "A1*2", excelize.FormulaOpts{Type: &formulaType, Ref: &ref}); err != nil {
		t.Fatal(err)
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	file.Close()

	ok(deleteRows(path, sheet, 8, 10))

	reopened, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()
	raw, ok2 := reopened.Pkg.Load("xl/worksheets/sheet1.xml")
	if !ok2 {
		t.Fatal("sheet1.xml missing from the package")
	}
	if !strings.Contains(string(raw.([]byte)), `<f t="shared" si="0"></f>`) {
		t.Errorf("shared group was expanded although nothing in it changed:\n%s", raw)
	}
}

func TestDeleteRowsMovesConditionalFormatsAndValidations(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "rules.xlsx")

	file := excelize.NewFile()
	sheet := file.GetSheetName(0)
	for row := 1; row <= 20; row++ {
		if err := file.SetCellInt(sheet, fmt.Sprintf("A%d", row), int64(row)); err != nil {
			t.Fatal(err)
		}
	}
	styleID, err := file.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "FF0000"}})
	if err != nil {
		t.Fatal(err)
	}
	if err := file.SetConditionalFormat(sheet, "A10:A20", []excelize.ConditionalFormatOptions{
		{Type: "cell", Criteria: ">", Format: &styleID, Value: "5"},
	}); err != nil {
		t.Fatal(err)
	}
	validation := excelize.NewDataValidation(true)
	validation.Sqref = "A10:A20"
	if err := validation.SetRange(1, 100, excelize.DataValidationTypeWhole, excelize.DataValidationOperatorBetween); err != nil {
		t.Fatal(err)
	}
	if err := file.AddDataValidation(sheet, validation); err != nil {
		t.Fatal(err)
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	file.Close()

	// Deleting rows 2-6 removes five rows above both rule ranges.
	output := ok(deleteRows(path, sheet, 2, 6))
	if !strings.Contains(output, "A5:A15") {
		t.Errorf("expected the shifted rule ranges to be reported, got: %s", output)
	}

	reopened, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()

	formats, err := reopened.GetConditionalFormats(sheet)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := formats["A5:A15"]; !ok {
		t.Errorf("conditional format did not shift to A5:A15, got %v", keysOf(formats))
	}
	validations, err := reopened.GetDataValidations(sheet)
	if err != nil {
		t.Fatal(err)
	}
	if len(validations) != 1 || validations[0].Sqref != "A5:A15" {
		t.Errorf("data validation did not shift to A5:A15, got %+v", validations)
	}
}

func TestInsertRowsShiftsContentDown(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "excel-mcp-tabtest.xlsx")
	buildTabTestWorkbook(t, path)

	ok(insertRows(path, "Blad3", 3, 2))

	file, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer file.Close()

	if got, _ := file.GetCellValue("Blad3", "A2"); got != "Username" {
		t.Errorf("A2 = %q, want the header to stay put", got)
	}
	if got, _ := file.GetCellValue("Blad3", "A3"); got != "" {
		t.Errorf("A3 = %q, want an empty inserted row", got)
	}
	if got, _ := file.GetCellValue("Blad3", "A5"); got != "oudklant1" {
		t.Errorf("A5 = %q, want oudklant1 shifted down by two", got)
	}
	if got, _ := file.GetCellValue("Blad3", "A10"); got != "actiefklant3" {
		t.Errorf("A10 = %q, want actiefklant3", got)
	}
	if got, _ := file.GetSheetDimension("Blad3"); got != "A2:D10" {
		t.Errorf("used range = %q, want A2:D10", got)
	}
}

// TestDeleteThenInsertRoundTrip covers the reason insert exists at all: a
// delete has to be undoable.
func TestDeleteThenInsertRoundTrip(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "excel-mcp-tabtest.xlsx")
	buildTabTestWorkbook(t, path)

	ok(deleteRows(path, "Blad3", 3, 5))
	ok(insertRows(path, "Blad3", 3, 3))

	file, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer file.Close()

	if got, _ := file.GetCellValue("Blad3", "A6"); got != "actiefklant1" {
		t.Errorf("A6 = %q, want actiefklant1 back at its original row", got)
	}
	if got, _ := file.GetSheetDimension("Blad3"); got != "A2:D8" {
		t.Errorf("used range = %q, want A2:D8", got)
	}
}

func TestRowToolsRejectInvalidRanges(t *testing.T) {
	fail := expectError(t)
	path := filepath.Join(t.TempDir(), "excel-mcp-tabtest.xlsx")
	buildTabTestWorkbook(t, path)

	message := fail(deleteRows(path, "Blad3", 5, 3))
	if !strings.Contains(message, "must not be smaller") {
		t.Errorf("expected a reversed-range error, got: %s", message)
	}
	message = fail(deleteRows(path, "Bestaat niet", 1, 1))
	if !strings.Contains(message, "sheet not found") {
		t.Errorf("expected a not-found error, got: %s", message)
	}
	message = fail(insertRows(path, "Bestaat niet", 1, 1))
	if !strings.Contains(message, "sheet not found") {
		t.Errorf("expected a not-found error, got: %s", message)
	}
}

// TestRowToolsAcceptAnyCasing pins the excelize quirk that makes FindSheet
// canonicalise the name: RemoveRow("data") on a sheet stored as "Data" moves
// the rows but leaves every Data!A1 reference in the workbook unshifted.
func TestRowToolsAcceptAnyCasing(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "casing.xlsx")

	file := excelize.NewFile()
	if err := file.SetSheetName(file.GetSheetName(0), "Data"); err != nil {
		t.Fatal(err)
	}
	for row := 1; row <= 20; row++ {
		if err := file.SetCellInt("Data", fmt.Sprintf("A%d", row), int64(row)); err != nil {
			t.Fatal(err)
		}
	}
	if _, err := file.NewSheet("Report"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellFormula("Report", "A1", "=Data!A20"); err != nil {
		t.Fatal(err)
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	file.Close()

	ok(deleteRows(path, "data", 5, 15))
	ok(insertRows(path, "DATA", 2, 3))

	reopened, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()
	if got, _ := reopened.GetCellFormula("Report", "A1"); got != "Data!A12" {
		t.Errorf("Report!A1 = %q, want Data!A12", got)
	}
	if got, _ := reopened.GetCellValue("Data", "A12"); got != "20" {
		t.Errorf("Data!A12 = %q, want 20", got)
	}
}

// TestDeleteRowsToTheBottom covers "delete everything below row N", which
// taken literally is a million excelize RemoveRow calls.
func TestDeleteRowsToTheBottom(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "excel-mcp-tabtest.xlsx")
	buildTabTestWorkbook(t, path)

	started := time.Now()
	ok(deleteRows(path, "Blad3", 4, excelize.TotalRows))
	if elapsed := time.Since(started); elapsed > 5*time.Second {
		t.Errorf("deleting to the bottom took %s, the loop is not capped at the used range", elapsed)
	}

	file, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer file.Close()
	if got, _ := file.GetCellValue("Blad3", "A3"); got != "oudklant1" {
		t.Errorf("A3 = %q, want oudklant1 to survive", got)
	}
	if got, _ := file.GetCellValue("Blad3", "A4"); got != "" {
		t.Errorf("A4 = %q, want everything from row 4 down gone", got)
	}
	if got, _ := file.GetSheetDimension("Blad3"); got != "A2:D3" {
		t.Errorf("used range = %q, want A2:D3", got)
	}
}

// TestReadSheetReportsRules covers the "am I about to destroy something I
// cannot see" question.
func TestReadSheetReportsRules(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "rules.xlsx")

	file := excelize.NewFile()
	sheet := file.GetSheetName(0)
	for row := 1; row <= 5; row++ {
		if err := file.SetCellInt(sheet, fmt.Sprintf("A%d", row), int64(row)); err != nil {
			t.Fatal(err)
		}
	}
	styleID, err := file.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "FF0000"}})
	if err != nil {
		t.Fatal(err)
	}
	if err := file.SetConditionalFormat(sheet, "A1:A5", []excelize.ConditionalFormatOptions{
		{Type: "cell", Criteria: ">", Format: &styleID, Value: "2"},
	}); err != nil {
		t.Fatal(err)
	}
	validation := excelize.NewDataValidation(true)
	validation.Sqref = "A1:A5"
	if err := validation.SetRange(1, 10, excelize.DataValidationTypeWhole, excelize.DataValidationOperatorBetween); err != nil {
		t.Fatal(err)
	}
	if err := file.AddDataValidation(sheet, validation); err != nil {
		t.Fatal(err)
	}
	if err := file.SetSheetDimension(sheet, "A1:A5"); err != nil {
		t.Fatal(err)
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	file.Close()

	output := ok(readSheet(path, sheet, "A1:A5", false, true))
	if !strings.Contains(output, "conditional formatting (1 rule range(s)): A1:A5") {
		t.Errorf("expected conditional formatting to be reported, got: %s", output)
	}
	if !strings.Contains(output, "data validation (1 rule range(s)): A1:A5") {
		t.Errorf("expected data validation to be reported, got: %s", output)
	}
}

func keysOf(m map[string][]excelize.ConditionalFormatOptions) []string {
	keys := make([]string, 0, len(m))
	for key := range m {
		keys = append(keys, key)
	}
	return keys
}
