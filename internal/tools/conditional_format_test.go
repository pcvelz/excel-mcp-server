package tools

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/negokaz/excel-mcp-server/internal/excel"
	"github.com/xuri/excelize/v2"
)

func newWorkbook(t *testing.T) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "book.xlsx")
	file := excelize.NewFile()
	for row := 1; row <= 5; row++ {
		if err := file.SetCellInt("Sheet1", cellRef(t, row), int64(row)); err != nil {
			t.Fatal(err)
		}
	}
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := file.Close(); err != nil {
		t.Fatal(err)
	}
	return path
}

func cellRef(t *testing.T, row int) string {
	t.Helper()
	ref, err := excelize.CoordinatesToCellName(1, row)
	if err != nil {
		t.Fatal(err)
	}
	return ref
}

func redFill() *excel.FillStyle {
	return &excel.FillStyle{Type: excel.FillTypePattern, Pattern: excel.FillPatternSolid, Color: []string{"#FFC7CE"}}
}

// Writing a rule has to survive the save and read back with the colour that
// was asked for. Asserting on the reader rather than the returned HTML is the
// point: it proves the rule is in the file, not just that the call succeeded.
func TestConditionalFormatWritesRuleThatReadsBack(t *testing.T) {
	path := newWorkbook(t)
	green := "#006100"

	output := expectOK(t)(conditionalFormat(path, "Sheet1", "A1:A5", "set", []*excel.ConditionalFormatRule{
		{Type: "cellIs", Operator: "greaterThan", Formulas: []string{"2"}, Fill: redFill(), StopIfTrue: true},
		{Type: "expression", Formulas: []string{"$A1=TODAY()"}, Font: &excel.FontStyle{Color: &green}},
	}))
	if !strings.Contains(output, "Applied 2 rule(s)") {
		t.Errorf("expected both rules to be reported as applied, got: %s", output)
	}

	rules := readBackRules(t, path)
	if len(rules) != 2 {
		t.Fatalf("expected 2 rules in the saved file, got %d: %+v", len(rules), rules)
	}
	var sawCellIs, sawExpression bool
	for _, rule := range rules {
		switch rule.Type {
		case "cellIs":
			sawCellIs = true
			if rule.Operator != "greaterThan" {
				t.Errorf("expected operator greaterThan, got %q", rule.Operator)
			}
			if len(rule.Formulas) != 1 || rule.Formulas[0] != "2" {
				t.Errorf("expected formula 2, got %v", rule.Formulas)
			}
			if rule.Fill == nil || len(rule.Fill.Color) == 0 || rule.Fill.Color[0] != "#FFC7CE" {
				t.Errorf("expected the fill colour to round trip, got %+v", rule.Fill)
			}
			if !rule.StopIfTrue {
				t.Error("expected stopIfTrue to round trip")
			}
		case "expression":
			sawExpression = true
			if len(rule.Formulas) != 1 || rule.Formulas[0] != "$A1=TODAY()" {
				t.Errorf("expected the custom formula to round trip, got %v", rule.Formulas)
			}
			if rule.Font == nil || rule.Font.Color == nil || *rule.Font.Color != green {
				t.Errorf("expected the font colour to round trip, got %+v", rule.Font)
			}
		}
		if rule.StyleError != "" {
			t.Errorf("the written style should resolve, got %q", rule.StyleError)
		}
	}
	if !sawCellIs || !sawExpression {
		t.Errorf("expected both rule types back, cellIs=%v expression=%v", sawCellIs, sawExpression)
	}
}

func TestConditionalFormatClearsARange(t *testing.T) {
	path := newWorkbook(t)
	expectOK(t)(conditionalFormat(path, "Sheet1", "A1:A5", "set", []*excel.ConditionalFormatRule{
		{Type: "cellIs", Operator: "greaterThan", Formulas: []string{"2"}, Fill: redFill()},
	}))
	if len(readBackRules(t, path)) != 1 {
		t.Fatal("expected the rule to be written before clearing")
	}

	output := expectOK(t)(conditionalFormat(path, "Sheet1", "A1:A5", "clear", nil))
	if !strings.Contains(output, "Cleared the conditional formatting on A1:A5") {
		t.Errorf("expected the clear to be reported, got: %s", output)
	}
	if rules := readBackRules(t, path); len(rules) != 0 {
		t.Errorf("expected no rules after clearing, got %+v", rules)
	}
}

// Writing must not quietly do nothing: a cellIs rule with no formula, or an
// operator the backend cannot express, has to come back as an error rather
// than a rule Excel ignores.
func TestConditionalFormatRejectsIncompleteRules(t *testing.T) {
	path := newWorkbook(t)
	for name, rule := range map[string]*excel.ConditionalFormatRule{
		"cellIs without a formula":     {Type: "cellIs", Operator: "greaterThan", Fill: redFill()},
		"between with one formula":     {Type: "cellIs", Operator: "between", Formulas: []string{"1"}, Fill: redFill()},
		"expression without a formula": {Type: "expression", Fill: redFill()},
	} {
		t.Run(name, func(t *testing.T) {
			result, err := conditionalFormat(path, "Sheet1", "A1:A5", "set", []*excel.ConditionalFormatRule{rule})
			if err != nil {
				t.Fatal(err)
			}
			if !result.IsError {
				t.Errorf("expected %s to be rejected", name)
			}
		})
	}

	result, err := conditionalFormat(path, "Sheet1", "A1:A5", "set", nil)
	if err != nil {
		t.Fatal(err)
	}
	if !result.IsError {
		t.Error("expected set with no rules to be rejected")
	}
}

func readBackRules(t *testing.T, path string) []excel.ConditionalFormatRule {
	t.Helper()
	book, release, err := excel.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer release()
	worksheet, err := book.FindSheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	rules, err := worksheet.GetConditionalFormats()
	if err != nil {
		t.Fatal(err)
	}
	return rules
}
