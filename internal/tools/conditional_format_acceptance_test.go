package tools

import (
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"

	"github.com/negokaz/excel-mcp-server/internal/excel"
)

// TestReadSheetConditionalFormattingAgainstRealWorkbook prints the
// conditional formatting line exactly as excel_read_sheet reports it, for a
// workbook of your choosing.
//
// No workbook ships with the repository: real ones carry real data. Point the
// test at your own file to run it:
//
//	EXCEL_MCP_ACCEPTANCE_WORKBOOK=/path/to/book.xlsx \
//	  go test ./internal/tools -run TestReadSheetConditionalFormattingAgainstRealWorkbook -v
//
// Set EXCEL_MCP_ACCEPTANCE_SHEET to pick a sheet, and
// EXCEL_MCP_ACCEPTANCE_RANGE to limit the read.
func TestReadSheetConditionalFormattingAgainstRealWorkbook(t *testing.T) {
	path := os.Getenv("EXCEL_MCP_ACCEPTANCE_WORKBOOK")
	if path == "" {
		t.Skip("set EXCEL_MCP_ACCEPTANCE_WORKBOOK to a workbook to run this")
	}
	sheetName := os.Getenv("EXCEL_MCP_ACCEPTANCE_SHEET")
	if sheetName == "" {
		sheetName = "Sheet1"
	}
	output := expectOK(t)(readSheet(path, sheetName, os.Getenv("EXCEL_MCP_ACCEPTANCE_RANGE"), false, true))

	line := regexp.MustCompile(`(?s)<li>conditional formatting.*?</li>`).FindString(output)
	if line == "" {
		t.Fatalf("no conditional formatting reported for %s!%s", path, sheetName)
	}
	t.Logf("%s", line)
	if regexp.MustCompile(`UNRESOLVED`).MatchString(line) {
		t.Errorf("a rule references a differential style the workbook does not define")
	}
}

// TestWritingDoesNotDisturbExistingRules works on a copy of your workbook and
// checks the thing that would actually hurt: adding a rule to one range must
// leave every other rule in the file untouched. Same environment variables as
// the test above; the original file is never written to.
func TestWritingDoesNotDisturbExistingRules(t *testing.T) {
	source := os.Getenv("EXCEL_MCP_ACCEPTANCE_WORKBOOK")
	if source == "" {
		t.Skip("set EXCEL_MCP_ACCEPTANCE_WORKBOOK to a workbook to run this")
	}
	sheetName := os.Getenv("EXCEL_MCP_ACCEPTANCE_SHEET")
	if sheetName == "" {
		sheetName = "Sheet1"
	}

	original, err := os.ReadFile(source)
	if err != nil {
		t.Fatal(err)
	}
	working := filepath.Join(t.TempDir(), "copy.xlsx")
	if err := os.WriteFile(working, original, 0o600); err != nil {
		t.Fatal(err)
	}

	before := readBackRulesOn(t, working, sheetName)
	t.Logf("workbook starts with %d rule(s)", len(before))

	fill := excel.FillStyle{Type: excel.FillTypePattern, Pattern: excel.FillPatternSolid, Color: []string{"#C6EFCE"}}
	expectOK(t)(conditionalFormat(working, sheetName, "Z1:Z5", "set", []*excel.ConditionalFormatRule{
		{Type: "cellIs", Operator: "greaterThan", Formulas: []string{"0"}, Fill: &fill},
	}))

	after := readBackRulesOn(t, working, sheetName)
	if len(after) != len(before)+1 {
		t.Fatalf("expected %d rules after adding one, got %d", len(before)+1, len(after))
	}
	// Every rule that was there has to still be there, unchanged.
	for _, want := range before {
		found := false
		for _, got := range after {
			if got.Range == want.Range && got.Type == want.Type && got.Operator == want.Operator &&
				strings.Join(got.Formulas, "\x00") == strings.Join(want.Formulas, "\x00") &&
				got.Priority == want.Priority && fillColorOf(got) == fillColorOf(want) {
				found = true
				break
			}
		}
		if !found {
			t.Errorf("a pre-existing rule was lost or altered: %+v", want)
		}
	}
}

func fillColorOf(rule excel.ConditionalFormatRule) string {
	if rule.Fill == nil || len(rule.Fill.Color) == 0 {
		return ""
	}
	return rule.Fill.Color[0]
}

func readBackRulesOn(t *testing.T, path, sheetName string) []excel.ConditionalFormatRule {
	t.Helper()
	book, release, err := excel.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer release()
	worksheet, err := book.FindSheet(sheetName)
	if err != nil {
		t.Fatal(err)
	}
	rules, err := worksheet.GetConditionalFormats()
	if err != nil {
		t.Fatal(err)
	}
	return rules
}
