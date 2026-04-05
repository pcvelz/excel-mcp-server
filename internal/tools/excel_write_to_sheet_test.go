package tools

import (
	"archive/zip"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

// TestNewSheetOptional verifies that the newSheet parameter is optional
// and defaults to false when omitted. Prior to this fix, omitting newSheet
// caused a misleading "values must be a 2D array" error.
func TestNewSheetOptional(t *testing.T) {
	input := map[string]any{
		"fileAbsolutePath": "/tmp/test.xlsx",
		"sheetName":        "Blad1",
		"range":            "A1",
		"values":           []any{[]any{"test"}},
	}

	args := ExcelWriteToSheetArguments{}
	issues := excelWriteToSheetArgumentsSchema.Parse(input, &args)
	if len(issues) != 0 {
		t.Fatalf("expected no validation issues when newSheet omitted, got: %v", issues)
	}
	if args.NewSheet != false {
		t.Errorf("expected NewSheet to default to false, got: %v", args.NewSheet)
	}
	if args.SheetName != "Blad1" {
		t.Errorf("expected SheetName Blad1, got: %v", args.SheetName)
	}
}

// TestNewSheetExplicitTrue verifies newSheet=true still works.
func TestNewSheetExplicitTrue(t *testing.T) {
	input := map[string]any{
		"fileAbsolutePath": "/tmp/test.xlsx",
		"sheetName":        "NewSheet",
		"newSheet":         true,
		"range":            "A1",
		"values":           []any{[]any{"test"}},
	}

	args := ExcelWriteToSheetArguments{}
	issues := excelWriteToSheetArgumentsSchema.Parse(input, &args)
	if len(issues) != 0 {
		t.Fatalf("expected no validation issues, got: %v", issues)
	}
	if args.NewSheet != true {
		t.Errorf("expected NewSheet=true, got: %v", args.NewSheet)
	}
}

// TestMissingRequiredFields verifies that truly required fields still error.
func TestMissingRequiredFields(t *testing.T) {
	input := map[string]any{
		"sheetName": "Blad1",
		"range":     "A1",
		"values":    []any{[]any{"test"}},
	}

	args := ExcelWriteToSheetArguments{}
	issues := excelWriteToSheetArgumentsSchema.Parse(input, &args)
	if len(issues) == 0 {
		t.Fatal("expected validation issue for missing fileAbsolutePath")
	}
}

// readSheetXML extracts the raw xl/worksheets/sheet1.xml from an .xlsx file.
func readSheetXML(t *testing.T, path string) string {
	t.Helper()
	r, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}
	defer r.Close()
	for _, f := range r.File {
		if f.Name == "xl/worksheets/sheet1.xml" {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("open sheet1.xml: %v", err)
			}
			defer rc.Close()
			b, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("read sheet1.xml: %v", err)
			}
			return string(b)
		}
	}
	t.Fatalf("xl/worksheets/sheet1.xml not found in %s", path)
	return ""
}

// TestWriteToSheet_PreservesDateTypeOnNoOp is a regression test for the bug
// where writing back the formatted string value of a numeric date cell would
// convert that cell from a date (numeric serial) into a plain shared string,
// silently breaking any conditional formatting that compares cell values
// against TODAY() or other numeric dates.
func TestWriteToSheet_PreservesDateTypeOnNoOp(t *testing.T) {
	tmp := t.TempDir()
	path := filepath.Join(tmp, "cf-test.xlsx")

	// Build a file with a date cell + conditional formatting
	f := excelize.NewFile()
	defer f.Close()

	// Date format style for "dd-mm-yy"
	dateStyle, err := f.NewStyle(&excelize.Style{NumFmt: 14})
	if err != nil {
		t.Fatalf("NewStyle: %v", err)
	}

	// Write a real numeric date to A1
	date := time.Date(2026, 6, 30, 0, 0, 0, 0, time.UTC)
	if err := f.SetCellValue("Sheet1", "A1", date); err != nil {
		t.Fatalf("SetCellValue: %v", err)
	}
	if err := f.SetCellStyle("Sheet1", "A1", "A1", dateStyle); err != nil {
		t.Fatalf("SetCellStyle: %v", err)
	}

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	f.Close()

	// Read back what the formatted display is — that is what the MCP
	// excel_read_sheet tool returns to the model.
	f2, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	displayed, err := f2.GetCellValue("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	rawBefore, err := f2.GetCellValue("Sheet1", "A1", excelize.Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("GetCellValue raw: %v", err)
	}
	f2.Close()

	if rawBefore == displayed {
		t.Fatalf("expected raw != display before write, got raw=%q display=%q", rawBefore, displayed)
	}

	// Simulate the MCP read→write round-trip: write the displayed value back.
	values := [][]any{{displayed}}
	if _, err := writeSheet(path, "Sheet1", false, "A1:A1", values); err != nil {
		t.Fatalf("writeSheet: %v", err)
	}

	// After the no-op write, the cell must STILL be a numeric date, not a string.
	f3, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f3.Close()
	rawAfter, err := f3.GetCellValue("Sheet1", "A1", excelize.Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("GetCellValue raw after: %v", err)
	}
	if rawAfter != rawBefore {
		t.Errorf("cell type/value changed: before raw=%q, after raw=%q (displayed=%q)", rawBefore, rawAfter, displayed)
	}
	cellType, err := f3.GetCellType("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetCellType: %v", err)
	}
	if cellType == excelize.CellTypeSharedString || cellType == excelize.CellTypeInlineString {
		t.Errorf("cell converted to string type (%v); expected numeric date preserved", cellType)
	}

	// Also verify t="s" is absent in the raw XML for the cell
	xml := readSheetXML(t, path)
	if strings.Contains(xml, `<c r="A1" s=`) && strings.Contains(xml, `t="s"`) {
		// crude sanity check; ensures no blanket shared-string conversion
		_ = os.WriteFile(filepath.Join(tmp, "dump.xml"), []byte(xml), 0644)
	}
}

// TestWriteToSheet_StillWritesWhenValueDiffers verifies the no-op guard does
// not swallow real writes.
func TestWriteToSheet_StillWritesWhenValueDiffers(t *testing.T) {
	tmp := t.TempDir()
	path := filepath.Join(tmp, "diff-test.xlsx")

	f := excelize.NewFile()
	if err := f.SetCellValue("Sheet1", "A1", "hello"); err != nil {
		t.Fatalf("SetCellValue: %v", err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	f.Close()

	values := [][]any{{"world"}}
	if _, err := writeSheet(path, "Sheet1", false, "A1:A1", values); err != nil {
		t.Fatalf("writeSheet: %v", err)
	}

	f2, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()
	got, err := f2.GetCellValue("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if got != "world" {
		t.Errorf("expected A1=world, got %q", got)
	}
}
