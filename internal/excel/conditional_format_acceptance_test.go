package excel

import (
	"os"
	"testing"
)

// TestConditionalFormatsAgainstRealWorkbook dumps the conditional formatting
// of a workbook of your choosing, so the reader can be checked against a file
// Excel itself produced rather than one the test suite built.
//
// No workbook ships with the repository: real ones carry real data. Point the
// test at your own file to run it:
//
//	EXCEL_MCP_ACCEPTANCE_WORKBOOK=/path/to/book.xlsx \
//	  go test ./internal/excel -run TestConditionalFormatsAgainstRealWorkbook -v
//
// Set EXCEL_MCP_ACCEPTANCE_SHEET to pick a sheet other than the first.
func TestConditionalFormatsAgainstRealWorkbook(t *testing.T) {
	path := os.Getenv("EXCEL_MCP_ACCEPTANCE_WORKBOOK")
	if path == "" {
		t.Skip("set EXCEL_MCP_ACCEPTANCE_WORKBOOK to a workbook to run this")
	}
	book, release, err := OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer release()

	sheetName := os.Getenv("EXCEL_MCP_ACCEPTANCE_SHEET")
	if sheetName == "" {
		names, err := book.SheetNames()
		if err != nil {
			t.Fatal(err)
		}
		if len(names) == 0 {
			t.Fatal("workbook has no sheets")
		}
		sheetName = names[0]
	}
	worksheet, err := book.FindSheet(sheetName)
	if err != nil {
		t.Fatal(err)
	}
	rules, err := worksheet.GetConditionalFormats()
	if err != nil {
		t.Fatal(err)
	}

	t.Logf("%s: %d conditional formatting rule(s), in priority order", sheetName, len(rules))
	for _, rule := range rules {
		fill, font := "-", "-"
		if rule.Fill != nil && len(rule.Fill.Color) > 0 {
			fill = rule.Fill.Color[0]
		}
		if rule.Font != nil && rule.Font.Color != nil {
			font = *rule.Font.Color
		}
		t.Logf("  priority=%-3d %-12s %-10s %-18s fill=%-9s font=%-9s stopIfTrue=%-5v formulas=%v colors=%v thresholds=%v",
			rule.Priority, rule.Range, rule.Type, rule.Operator, fill, font,
			rule.StopIfTrue, rule.Formulas, rule.Colors, rule.Thresholds)
	}
}
