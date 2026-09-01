package tools

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/mark3labs/mcp-go/mcp"
	"github.com/xuri/excelize/v2"
)

// expectOK returns a checker that flattens a successful tool result into text.
// It is a closure so that call sites can pass a tool call straight through:
// Go only allows f(g()) when g's results are f's only arguments.
func expectOK(t *testing.T) func(*mcp.CallToolResult, error) string {
	return func(result *mcp.CallToolResult, err error) string {
		t.Helper()
		text := flattenResult(result)
		if err != nil {
			t.Fatalf("tool returned an error: %v", err)
		}
		if result.IsError {
			t.Fatalf("tool reported failure: %s", text)
		}
		return text
	}
}

// expectError returns a checker for tool calls that must fail.
func expectError(t *testing.T) func(*mcp.CallToolResult, error) string {
	return func(result *mcp.CallToolResult, err error) string {
		t.Helper()
		if err != nil {
			t.Fatalf("tool returned an unexpected Go error: %v", err)
		}
		text := flattenResult(result)
		if !result.IsError {
			t.Fatalf("expected the tool to report failure, got: %s", text)
		}
		return text
	}
}

func flattenResult(result *mcp.CallToolResult) string {
	var builder strings.Builder
	for _, content := range result.Content {
		if text, ok := content.(mcp.TextContent); ok {
			builder.WriteString(text.Text)
		}
	}
	return builder.String()
}

// buildTabTestWorkbook writes the fixture used by the sheet management tests:
// three sheets whose formatting carries meaning, so that any silent formatting
// loss shows up as a test failure.
func buildTabTestWorkbook(t *testing.T, path string) {
	t.Helper()

	file := excelize.NewFile()
	defer file.Close()

	if err := file.SetSheetName(file.GetSheetName(0), "Blad3"); err != nil {
		t.Fatal(err)
	}

	headers := []string{"Username", "Bedrijfsnaam", "E-mail", "Betaald"}
	for i, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 2)
		if err := file.SetCellStr("Blad3", cell, header); err != nil {
			t.Fatal(err)
		}
	}

	type client struct {
		username string
		company  string
		email    string
		paid     string
	}
	former := []client{
		{"oudklant1", "Beeindigde zaak 1", "oud1@example.invalid", "2019-01-01"},
		{"oudklant2", "Beeindigde zaak 2", "oud2@example.invalid", "2020-10-31"},
		{"oudklant3", "Beeindigde zaak 3", "oud3@example.invalid", "2024-04-30"},
	}
	active := []client{
		{"actiefklant1", "Lopende zaak 1", "a@example.invalid", "2026-10-31"},
		{"actiefklant2", "Lopende zaak 2", "b@example.invalid", "2026-10-31"},
		{"actiefklant3", "Lopende zaak 3", "c@example.invalid", "2026-12-31"},
	}

	writeRow := func(row int, c client) {
		if err := file.SetCellStr("Blad3", fmt.Sprintf("A%d", row), c.username); err != nil {
			t.Fatal(err)
		}
		if err := file.SetCellStr("Blad3", fmt.Sprintf("B%d", row), c.company); err != nil {
			t.Fatal(err)
		}
		if err := file.SetCellStr("Blad3", fmt.Sprintf("C%d", row), c.email); err != nil {
			t.Fatal(err)
		}
		if err := file.SetCellStr("Blad3", fmt.Sprintf("D%d", row), c.paid); err != nil {
			t.Fatal(err)
		}
	}
	for i, c := range former {
		writeRow(3+i, c)
	}
	for i, c := range active {
		writeRow(6+i, c)
	}

	// Font colour is meaningful here: green marks a client that is still
	// active. Rows 2-8 are bold at size 11 throughout.
	blackStyle, err := file.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Size: 11, Color: "000000"},
	})
	if err != nil {
		t.Fatal(err)
	}
	greenStyle, err := file.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Size: 11, Color: "00B050"},
	})
	if err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellStyle("Blad3", "A2", "D5", blackStyle); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellStyle("Blad3", "A6", "D8", greenStyle); err != nil {
		t.Fatal(err)
	}

	// D2 wraps its text and A4 is vertically centred, so alignment loss is
	// visible too.
	wrapStyle, err := file.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 11, Color: "000000"},
		Alignment: &excelize.Alignment{WrapText: true},
	})
	if err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellStyle("Blad3", "D2", "D2", wrapStyle); err != nil {
		t.Fatal(err)
	}
	middleStyle, err := file.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 11, Color: "000000"},
		Alignment: &excelize.Alignment{Vertical: "center"},
	})
	if err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellStyle("Blad3", "A4", "A4", middleStyle); err != nil {
		t.Fatal(err)
	}

	if err := file.MergeCell("Blad3", "B2", "C2"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetColWidth("Blad3", "B", "B", 40); err != nil {
		t.Fatal(err)
	}
	// excelize does not maintain the <dimension> element when cells are written
	// directly, and the read tool paginates on it, so set it explicitly.
	if err := file.SetSheetDimension("Blad3", "A2:D8"); err != nil {
		t.Fatal(err)
	}

	if _, err := file.NewSheet("Blad1"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellStr("Blad1", "D6", "dit blad moet verwijderd worden"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetSheetDimension("Blad1", "D6:D6"); err != nil {
		t.Fatal(err)
	}

	if _, err := file.NewSheet("Klanten actief"); err != nil {
		t.Fatal(err)
	}
	activeHeaders := []string{"UUID", "Username", "Bedrijfsnaam"}
	for i, header := range activeHeaders {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		if err := file.SetCellStr("Klanten actief", cell, header); err != nil {
			t.Fatal(err)
		}
	}
	uuids := []string{
		"6f1a2c3d-4e5f-4a6b-8c9d-0e1f2a3b4c5d",
		"7a2b3c4d-5e6f-4b7c-9d0e-1f2a3b4c5d6e",
		"8b3c4d5e-6f70-4c8d-ae1f-2a3b4c5d6e7f",
	}
	for i, c := range active {
		row := 2 + i
		if err := file.SetCellStr("Klanten actief", fmt.Sprintf("A%d", row), uuids[i]); err != nil {
			t.Fatal(err)
		}
		if err := file.SetCellStr("Klanten actief", fmt.Sprintf("B%d", row), c.username); err != nil {
			t.Fatal(err)
		}
		if err := file.SetCellStr("Klanten actief", fmt.Sprintf("C%d", row), c.company); err != nil {
			t.Fatal(err)
		}
	}

	if err := file.SetSheetDimension("Klanten actief", "A1:C4"); err != nil {
		t.Fatal(err)
	}

	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
}

// TestSheetManagementAcceptance runs rename, delete and move end to end and
// then asserts that none of the formatting on the renamed sheet was lost.
func TestSheetManagementAcceptance(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "excel-mcp-tabtest.xlsx")
	buildTabTestWorkbook(t, path)

	ok(renameSheet(path, "Blad3", "Archief klanten"))
	ok(deleteSheet(path, "Blad1", false))
	ok(moveSheet(path, "Klanten actief", 0))

	file, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer file.Close()

	if got, want := file.GetSheetList(), []string{"Klanten actief", "Archief klanten"}; !equalStrings(got, want) {
		t.Fatalf("sheet order = %v, want %v", got, want)
	}

	sheet := "Archief klanten"
	if got, _ := file.GetCellValue(sheet, "A3"); got != "oudklant1" {
		t.Errorf("A3 = %q, want oudklant1", got)
	}
	if got, _ := file.GetCellValue(sheet, "A8"); got != "actiefklant3" {
		t.Errorf("A8 = %q, want actiefklant3", got)
	}
	if got, _ := file.GetSheetDimension(sheet); got != "A2:D8" {
		t.Errorf("used range = %q, want A2:D8", got)
	}

	// Font colour: rows 6-8 green, rows 3-5 not green, everything bold.
	for row := 2; row <= 8; row++ {
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
			if style.Font == nil || !style.Font.Bold {
				t.Errorf("%s is not bold", cell)
				continue
			}
			if style.Font.Size != 11 {
				t.Errorf("%s font size = %v, want 11", cell, style.Font.Size)
			}
			green := strings.Contains(strings.ToUpper(style.Font.Color), "00B050")
			if row >= 6 && !green {
				t.Errorf("%s font colour = %q, want 00B050", cell, style.Font.Color)
			}
			if row < 6 && green {
				t.Errorf("%s font colour = %q, want a non-green colour", cell, style.Font.Color)
			}
		}
	}

	// Alignment.
	styleID, _ := file.GetCellStyle(sheet, "D2")
	style, err := file.GetStyle(styleID)
	if err != nil {
		t.Fatal(err)
	}
	if style.Alignment == nil || !style.Alignment.WrapText {
		t.Error("D2 lost its wrapText alignment")
	}
	styleID, _ = file.GetCellStyle(sheet, "A4")
	style, err = file.GetStyle(styleID)
	if err != nil {
		t.Fatal(err)
	}
	if style.Alignment == nil || style.Alignment.Vertical != "center" {
		t.Error("A4 lost its vertical centre alignment")
	}

	// Merged cells and column width.
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

	// The untouched sheet must have survived the reordering intact.
	if got, _ := file.GetCellValue("Klanten actief", "A2"); len(got) != 36 {
		t.Errorf("Klanten actief A2 = %q, want a 36 character uuid", got)
	}
	if got, _ := file.GetCellValue("Klanten actief", "B4"); got != "actiefklant3" {
		t.Errorf("Klanten actief B4 = %q, want actiefklant3", got)
	}
}

// TestRenameSheetUpdatesFormulas covers the case excelize does not handle:
// formulas that point at the renamed sheet.
func TestRenameSheetUpdatesFormulas(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "formulas.xlsx")
	file := excelize.NewFile()
	if err := file.SetSheetName(file.GetSheetName(0), "Data"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellInt("Data", "A1", 21); err != nil {
		t.Fatal(err)
	}
	if _, err := file.NewSheet("Report"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellFormula("Report", "A1", "=Data!A1*2"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellStr("Report", "B1", "label"); err != nil {
		t.Fatal(err)
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	file.Close()

	ok(renameSheet(path, "Data", "Bron gegevens"))

	reopened, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()

	formula, err := reopened.GetCellFormula("Report", "A1")
	if err != nil {
		t.Fatal(err)
	}
	if formula != "='Bron gegevens'!A1*2" {
		t.Errorf("formula = %q, want ='Bron gegevens'!A1*2", formula)
	}
	value, err := reopened.CalcCellValue("Report", "A1")
	if err != nil {
		t.Fatal(err)
	}
	if value != "42" {
		t.Errorf("calculated value = %q, want 42", value)
	}
}

// TestRenameSheetKeepsSharedFormulaGroups guards against writing formula text
// into the members of a shared formula group: only the master carries text,
// and Excel would otherwise read the unshifted master formula in every member.
func TestRenameSheetKeepsSharedFormulaGroups(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "shared.xlsx")
	file := excelize.NewFile()
	if err := file.SetSheetName(file.GetSheetName(0), "Data"); err != nil {
		t.Fatal(err)
	}
	for row := 1; row <= 3; row++ {
		if err := file.SetCellInt("Data", fmt.Sprintf("A%d", row), int64(row)); err != nil {
			t.Fatal(err)
		}
	}
	if _, err := file.NewSheet("Report"); err != nil {
		t.Fatal(err)
	}
	formulaType, ref := excelize.STCellFormulaTypeShared, "C1:C3"
	if err := file.SetCellFormula("Report", "C1", "Data!A1*2", excelize.FormulaOpts{Type: &formulaType, Ref: &ref}); err != nil {
		t.Fatal(err)
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	file.Close()

	ok(renameSheet(path, "Data", "Bron"))

	reopened, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()
	raw, found := reopened.Pkg.Load("xl/worksheets/sheet2.xml")
	if !found {
		t.Fatal("sheet2.xml missing from the package")
	}
	xml := string(raw.([]byte))
	if !strings.Contains(xml, `<f t="shared" ref="C1:C3" si="0">Bron!A1*2</f>`) {
		t.Errorf("master formula was not rewritten in place:\n%s", xml)
	}
	if strings.Count(xml, "Bron!A1*2") != 1 || strings.Contains(xml, "Data!") {
		t.Errorf("group members must inherit the master rather than carry their own text:\n%s", xml)
	}
}

func TestRenameSheetRejectsDuplicateAndInvalidNames(t *testing.T) {
	fail := expectError(t)
	path := filepath.Join(t.TempDir(), "excel-mcp-tabtest.xlsx")
	buildTabTestWorkbook(t, path)

	message := fail(renameSheet(path, "Blad3", "Blad1"))
	if !strings.Contains(message, "already exists") {
		t.Errorf("expected a duplicate-name error, got: %s", message)
	}
	message = fail(renameSheet(path, "Blad3", "Naam/met/slash"))
	if !strings.Contains(message, "does not allow") {
		t.Errorf("expected an invalid-character error, got: %s", message)
	}
	message = fail(renameSheet(path, "Blad3", strings.Repeat("x", 32)))
	if !strings.Contains(message, "at most 31") {
		t.Errorf("expected a length error, got: %s", message)
	}
	message = fail(renameSheet(path, "Bestaat niet", "Nieuw"))
	if !strings.Contains(message, "sheet not found") {
		t.Errorf("expected a not-found error, got: %s", message)
	}
}

func TestDeleteSheetRefusesLastSheet(t *testing.T) {
	fail := expectError(t)
	path := filepath.Join(t.TempDir(), "single.xlsx")
	file := excelize.NewFile()
	if err := file.SetCellStr(file.GetSheetName(0), "A1", "only"); err != nil {
		t.Fatal(err)
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	name := file.GetSheetName(0)
	file.Close()

	message := fail(deleteSheet(path, name, false))
	if !strings.Contains(message, "at least one sheet") {
		t.Errorf("expected a last-sheet error, got: %s", message)
	}
}

func TestDeleteSheetRefusesWhenFormulasReferenceIt(t *testing.T) {
	ok := expectOK(t)
	fail := expectError(t)
	path := filepath.Join(t.TempDir(), "referenced.xlsx")
	file := excelize.NewFile()
	if err := file.SetSheetName(file.GetSheetName(0), "Data"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellInt("Data", "A1", 7); err != nil {
		t.Fatal(err)
	}
	if _, err := file.NewSheet("Report"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellFormula("Report", "A1", "=Data!A1"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellStr("Report", "B1", "label"); err != nil {
		t.Fatal(err)
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	file.Close()

	message := fail(deleteSheet(path, "Data", false))
	if !strings.Contains(message, "Report!A1") {
		t.Errorf("expected the referencing cell to be named, got: %s", message)
	}

	// force=true goes ahead and says what it broke.
	output := ok(deleteSheet(path, "Data", true))
	if !strings.Contains(output, "Warnings") || !strings.Contains(output, "Report!A1") {
		t.Errorf("expected a warning about the broken formula, got: %s", output)
	}
}

func TestDeleteSheetRemovesDefinedNames(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "definednames.xlsx")
	file := excelize.NewFile()
	if err := file.SetSheetName(file.GetSheetName(0), "Data"); err != nil {
		t.Fatal(err)
	}
	if _, err := file.NewSheet("Keep"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetCellStr("Keep", "A1", "keep"); err != nil {
		t.Fatal(err)
	}
	if err := file.SetDefinedName(&excelize.DefinedName{
		Name:     "Bereik",
		RefersTo: "Data!$A$1:$A$9",
		Scope:    "Workbook",
	}); err != nil {
		t.Fatal(err)
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	file.Close()

	output := ok(deleteSheet(path, "Data", false))
	if !strings.Contains(output, "Bereik") {
		t.Errorf("expected the removed defined name to be reported, got: %s", output)
	}

	reopened, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer reopened.Close()
	for _, definedName := range reopened.GetDefinedName() {
		if definedName.Name == "Bereik" {
			t.Errorf("defined name Bereik still refers to the deleted sheet: %+v", definedName)
		}
	}
}

func TestMoveSheetToEveryPosition(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "order.xlsx")
	file := excelize.NewFile()
	if err := file.SetSheetName(file.GetSheetName(0), "A"); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"B", "C", "D"} {
		if _, err := file.NewSheet(name); err != nil {
			t.Fatal(err)
		}
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	file.Close()

	tests := []struct {
		sheet string
		index int
		want  []string
	}{
		{"D", 0, []string{"D", "A", "B", "C"}},
		{"D", 3, []string{"A", "B", "C", "D"}},
		{"A", 2, []string{"B", "C", "A", "D"}},
		{"A", 0, []string{"A", "B", "C", "D"}},
		{"B", 3, []string{"A", "C", "D", "B"}},
	}
	for _, test := range tests {
		ok(moveSheet(path, test.sheet, test.index))
		reopened, err := excelize.OpenFile(path)
		if err != nil {
			t.Fatal(err)
		}
		got := reopened.GetSheetList()
		reopened.Close()
		if !equalStrings(got, test.want) {
			t.Fatalf("after moving %s to %d: order = %v, want %v", test.sheet, test.index, got, test.want)
		}
		// Reset for the next case.
		file, err := excelize.OpenFile(path)
		if err != nil {
			t.Fatal(err)
		}
		for i, name := range []string{"A", "B", "C", "D"} {
			if i == 0 {
				continue
			}
			if err := file.MoveSheet(name, []string{"A", "B", "C", "D"}[i-1]); err != nil {
				t.Fatal(err)
			}
		}
		file.Close()
	}
}

func TestMoveSheetRejectsOutOfRangeIndex(t *testing.T) {
	fail := expectError(t)
	path := filepath.Join(t.TempDir(), "excel-mcp-tabtest.xlsx")
	buildTabTestWorkbook(t, path)

	message := fail(moveSheet(path, "Blad1", 3))
	if !strings.Contains(message, "out of range") {
		t.Errorf("expected an out-of-range error, got: %s", message)
	}
}

// TestWriteToSheetCreatesNewWorkbook covers the reported bug: newSheet on a
// path that does not exist yet had no way to create the file.
func TestWriteToSheetCreatesNewWorkbook(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "brand-new.xlsx")

	ok(writeSheet(path, "Overzicht", true, "A1:B2", [][]any{
		{"Naam", "Bedrag"},
		{"Klant", "100"},
	}))

	file, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("expected the workbook to have been created: %v", err)
	}
	defer file.Close()

	if got, want := file.GetSheetList(), []string{"Overzicht"}; !equalStrings(got, want) {
		t.Errorf("sheet list = %v, want %v", got, want)
	}
	if got, _ := file.GetCellValue("Overzicht", "A1"); got != "Naam" {
		t.Errorf("A1 = %q, want Naam", got)
	}
	if got, _ := file.GetCellValue("Overzicht", "B2"); got != "100" {
		t.Errorf("B2 = %q, want 100", got)
	}
}

func TestWriteToSheetRejectsUncreatablePaths(t *testing.T) {
	fail := expectError(t)
	message := fail(writeSheet(filepath.Join(t.TempDir(), "notes.txt"), "Blad1", true, "A1", [][]any{{"x"}}))
	if !strings.Contains(message, "unsupported file extension") {
		t.Errorf("expected an extension error, got: %s", message)
	}
	message = fail(writeSheet(filepath.Join(t.TempDir(), "geen", "map", "boek.xlsx"), "Blad1", true, "A1", [][]any{{"x"}}))
	if !strings.Contains(message, "does not exist") {
		t.Errorf("expected a missing-directory error, got: %s", message)
	}
}

// TestReadSheetReportsMergesAndWidths checks that the facts a rename could
// silently destroy are actually observable through the read tool.
func TestReadSheetReportsMergesAndWidths(t *testing.T) {
	ok := expectOK(t)
	path := filepath.Join(t.TempDir(), "excel-mcp-tabtest.xlsx")
	buildTabTestWorkbook(t, path)

	output := ok(readSheet(path, "Blad3", "A2:D8", false, true))
	if !strings.Contains(output, "merged cells: B2:C2") {
		t.Errorf("expected the merge to be reported, got: %s", output)
	}
	if !strings.Contains(output, "B=40") {
		t.Errorf("expected column B width 40 to be reported, got: %s", output)
	}
	if !strings.Contains(output, "used range: A2:D8") {
		t.Errorf("expected the used range to be reported, got: %s", output)
	}
}

func equalStrings(a []string, b []string) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

// TestWriteTabTestFixture writes the shared fixture to the path named by
// EXCEL_MCP_TABTEST_PATH so it can be exercised by hand, keeping a single
// definition of the fixture. Skipped unless that variable is set.
func TestWriteTabTestFixture(t *testing.T) {
	path := os.Getenv("EXCEL_MCP_TABTEST_PATH")
	if path == "" {
		t.Skip("set EXCEL_MCP_TABTEST_PATH to write the fixture to disk")
	}
	buildTabTestWorkbook(t, path)
	t.Logf("wrote fixture to %s", path)
}
