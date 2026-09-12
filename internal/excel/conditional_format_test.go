package excel

import (
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

// openFixture builds a workbook, saves it, and reopens it through the package
// under test. Reopening matters: the rules have to survive a round trip
// through the XML, which is where they are actually read from.
func openFixture(t *testing.T, build func(*excelize.File)) Worksheet {
	t.Helper()
	path := filepath.Join(t.TempDir(), "conditional.xlsx")
	file := excelize.NewFile()
	build(file)
	if err := file.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := file.Close(); err != nil {
		t.Fatal(err)
	}
	book, release, err := OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(release)
	worksheet, err := book.FindSheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	return worksheet
}

func conditionalStyle(t *testing.T, file *excelize.File, fill string) *int {
	t.Helper()
	id, err := file.NewConditionalStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{fill}},
	})
	if err != nil {
		t.Fatal(err)
	}
	return &id
}

// Excel routinely writes several <conditionalFormatting> blocks that share one
// sqref -- adding a rule to a range that already has one produces exactly
// that. Keying rules by range collapses them, so this guards the shape of the
// result, not just its contents.
func TestConditionalFormatsKeepsEveryRuleOnARepeatedRange(t *testing.T) {
	worksheet := openFixture(t, func(file *excelize.File) {
		for _, rule := range []excelize.ConditionalFormatOptions{
			{Type: "cell", Criteria: "<=", Format: conditionalStyle(t, file, "FFC7CE"), Value: "10"},
			{Type: "cell", Criteria: ">", Format: conditionalStyle(t, file, "C6EFCE"), Value: "10"},
		} {
			// One call per rule, so each lands in its own block under the
			// same sqref.
			if err := file.SetConditionalFormat("Sheet1", "A1:A5", []excelize.ConditionalFormatOptions{rule}); err != nil {
				t.Fatal(err)
			}
		}
	})

	rules, err := worksheet.GetConditionalFormats()
	if err != nil {
		t.Fatal(err)
	}
	if len(rules) != 2 {
		t.Fatalf("expected both rules on A1:A5 to survive, got %d: %+v", len(rules), rules)
	}
	fills := map[string]bool{}
	for _, rule := range rules {
		if rule.Range != "A1:A5" {
			t.Errorf("expected range A1:A5, got %q", rule.Range)
		}
		if rule.Fill == nil || len(rule.Fill.Color) == 0 {
			t.Fatalf("rule has no fill colour: %+v", rule)
		}
		fills[rule.Fill.Color[0]] = true
	}
	for _, want := range []string{"#FFC7CE", "#C6EFCE"} {
		if !fills[want] {
			t.Errorf("expected a rule filling %s, got %v", want, fills)
		}
	}
}

// Where ranges overlap, priority alone decides what a cell ends up looking
// like, so the rules have to come back in that order.
func TestConditionalFormatsAreOrderedByPriority(t *testing.T) {
	worksheet := openFixture(t, func(file *excelize.File) {
		if err := file.SetConditionalFormat("Sheet1", "A1:A10", []excelize.ConditionalFormatOptions{
			{Type: "cell", Criteria: "<=", Format: conditionalStyle(t, file, "FFC7CE"), Value: "1"},
			{Type: "cell", Criteria: "<=", Format: conditionalStyle(t, file, "FFEB9C"), Value: "2"},
		}); err != nil {
			t.Fatal(err)
		}
		if err := file.SetConditionalFormat("Sheet1", "A9:A10", []excelize.ConditionalFormatOptions{
			{Type: "cell", Criteria: ">", Format: conditionalStyle(t, file, "C6EFCE"), Value: "3"},
		}); err != nil {
			t.Fatal(err)
		}
	})

	rules, err := worksheet.GetConditionalFormats()
	if err != nil {
		t.Fatal(err)
	}
	if len(rules) != 3 {
		t.Fatalf("expected 3 rules, got %d", len(rules))
	}
	for i := 1; i < len(rules); i++ {
		if rules[i-1].Priority > rules[i].Priority {
			t.Errorf("rules are not in priority order: %d before %d", rules[i-1].Priority, rules[i].Priority)
		}
	}
}

// The formula and the operator live in separate fields and must not be
// conflated: for a cellIs rule the operator is the comparison and the formula
// is what it compares against.
func TestConditionalFormatsSeparatesOperatorFromFormula(t *testing.T) {
	worksheet := openFixture(t, func(file *excelize.File) {
		if err := file.SetConditionalFormat("Sheet1", "B1:B5", []excelize.ConditionalFormatOptions{
			{Type: "cell", Criteria: "between", Format: conditionalStyle(t, file, "FFEB9C"), MinValue: "TODAY()", MaxValue: "TODAY() + 30"},
		}); err != nil {
			t.Fatal(err)
		}
	})

	rules, err := worksheet.GetConditionalFormats()
	if err != nil {
		t.Fatal(err)
	}
	if len(rules) != 1 {
		t.Fatalf("expected 1 rule, got %d", len(rules))
	}
	rule := rules[0]
	if rule.Type != "cellIs" {
		t.Errorf("expected type cellIs, got %q", rule.Type)
	}
	if rule.Operator != "between" {
		t.Errorf("expected operator between, got %q", rule.Operator)
	}
	if len(rule.Formulas) != 2 || rule.Formulas[0] != "TODAY()" || rule.Formulas[1] != "TODAY() + 30" {
		t.Errorf("expected both bounds as formulas, got %v", rule.Formulas)
	}
}

// A custom-formula rule is the most common way to colour a row by the value of
// another cell, and stores its formula where a cellIs rule stores nothing.
func TestConditionalFormatsReadsCustomFormulaRules(t *testing.T) {
	worksheet := openFixture(t, func(file *excelize.File) {
		if err := file.SetConditionalFormat("Sheet1", "A1:C5", []excelize.ConditionalFormatOptions{
			{Type: "formula", Format: conditionalStyle(t, file, "C6EFCE"), Criteria: "$C1>TODAY()"},
		}); err != nil {
			t.Fatal(err)
		}
	})

	rules, err := worksheet.GetConditionalFormats()
	if err != nil {
		t.Fatal(err)
	}
	if len(rules) != 1 {
		t.Fatalf("expected 1 rule, got %d", len(rules))
	}
	if rules[0].Type != "expression" {
		t.Errorf("expected type expression, got %q", rules[0].Type)
	}
	if len(rules[0].Formulas) != 1 || rules[0].Formulas[0] != "$C1>TODAY()" {
		t.Errorf("expected the custom formula, got %v", rules[0].Formulas)
	}
}

// A rule whose dxfId names a style the workbook never defines renders as no
// formatting in Excel. Reading such a sheet has to keep working and say what
// is wrong, rather than failing the whole read or inventing a colour.
func TestConditionalFormatsReportsUnresolvableStyle(t *testing.T) {
	worksheet := openFixture(t, func(file *excelize.File) {
		// NewStyle indexes cellXfs, not dxfs, so this leaves a dangling
		// dxfId -- the exact mistake that makes a rule invisible in Excel.
		wrongTable, err := file.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "FF0000"}})
		if err != nil {
			t.Fatal(err)
		}
		if err := file.SetConditionalFormat("Sheet1", "A1:A5", []excelize.ConditionalFormatOptions{
			{Type: "cell", Criteria: ">", Format: &wrongTable, Value: "2"},
		}); err != nil {
			t.Fatal(err)
		}
	})

	rules, err := worksheet.GetConditionalFormats()
	if err != nil {
		t.Fatalf("one malformed rule must not fail the read: %v", err)
	}
	if len(rules) != 1 {
		t.Fatalf("expected the rule to still be reported, got %d", len(rules))
	}
	if rules[0].StyleError == "" {
		t.Errorf("expected the dangling differential style to be reported, got %+v", rules[0])
	}
	if rules[0].Fill != nil {
		t.Errorf("expected no fill for an unresolvable style, got %+v", rules[0].Fill)
	}
}
