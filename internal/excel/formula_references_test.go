package excel

import (
	"strconv"
	"strings"
	"testing"
)

func TestReplaceSheetNameInFormula(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		old     string
		new     string
		want    string
		changed bool
	}{
		{
			name:    "bare reference",
			formula: "=Blad3!A1",
			old:     "Blad3",
			new:     "Archief",
			want:    "=Archief!A1",
			changed: true,
		},
		{
			name:    "new name with a space gets quoted",
			formula: "=Blad3!A1+Blad3!B2",
			old:     "Blad3",
			new:     "Archief klanten",
			want:    "='Archief klanten'!A1+'Archief klanten'!B2",
			changed: true,
		},
		{
			name:    "quoted reference",
			formula: "=SUM('Oude klanten'!A1:A9)",
			old:     "Oude klanten",
			new:     "Archief",
			want:    "=SUM(Archief!A1:A9)",
			changed: true,
		},
		{
			name:    "quote inside the old name",
			formula: "=SUM('Pete''s blad'!A1)",
			old:     "Pete's blad",
			new:     "Archief",
			want:    "=SUM(Archief!A1)",
			changed: true,
		},
		{
			name:    "quote inside the new name is escaped",
			formula: "=Blad3!A1",
			old:     "Blad3",
			new:     "Pete's blad",
			want:    "='Pete''s blad'!A1",
			changed: true,
		},
		{
			name:    "string literal is not touched",
			formula: `=IF(A1="Blad3!",Blad3!B1,"")`,
			old:     "Blad3",
			new:     "Archief",
			want:    `=IF(A1="Blad3!",Archief!B1,"")`,
			changed: true,
		},
		{
			name:    "escaped quote inside a string literal",
			formula: `=CONCAT("a""Blad3!b",Blad3!C1)`,
			old:     "Blad3",
			new:     "Archief",
			want:    `=CONCAT("a""Blad3!b",Archief!C1)`,
			changed: true,
		},
		{
			name:    "external workbook reference is left alone",
			formula: "=[1]Blad3!A1",
			old:     "Blad3",
			new:     "Archief",
			want:    "=[1]Blad3!A1",
			changed: false,
		},
		{
			name:    "function name that merely contains the sheet name",
			formula: "=Blad30!A1",
			old:     "Blad3",
			new:     "Archief",
			want:    "=Blad30!A1",
			changed: false,
		},
		{
			name:    "defined name without a bang is not a sheet reference",
			formula: "=Blad3",
			old:     "Blad3",
			new:     "Archief",
			want:    "=Blad3",
			changed: false,
		},
		{
			name:    "sheet name matched case-insensitively",
			formula: "=blad3!A1",
			old:     "Blad3",
			new:     "Archief",
			want:    "=Archief!A1",
			changed: true,
		},
		{
			name:    "new name that looks like a cell reference gets quoted",
			formula: "=Blad3!A1",
			old:     "Blad3",
			new:     "A1",
			want:    "='A1'!A1",
			changed: true,
		},
		{
			name:    "unrelated formula",
			formula: "=SUM(A1:A9)",
			old:     "Blad3",
			new:     "Archief",
			want:    "=SUM(A1:A9)",
			changed: false,
		},
		{
			name:    "broken reference keeps its qualifier",
			formula: "=Blad3!#REF!+1",
			old:     "Blad3",
			new:     "Archief",
			want:    "=Archief!#REF!+1",
			changed: true,
		},
		{
			name:    "error literal is not a sheet named REF",
			formula: "=#REF!+A1",
			old:     "REF",
			new:     "Archief",
			want:    "=#REF!+A1",
			changed: false,
		},
		{
			name:    "range with a function as its second half",
			formula: "=SUM(Blad3!A1:INDEX(Blad3!A:A,3))",
			old:     "Blad3",
			new:     "Archief",
			want:    "=SUM(Archief!A1:INDEX(Archief!A:A,3))",
			changed: true,
		},
		{
			name:    "sheet-qualified defined name is renamed like any other reference",
			formula: "=Blad3!MyName",
			old:     "Blad3",
			new:     "Archief",
			want:    "=Archief!MyName",
			changed: true,
		},
		{
			name:    "string literal containing a cell reference is not touched",
			formula: `=IF(A1="Blad3!A6","x","y")`,
			old:     "Blad3",
			new:     "Archief",
			want:    `=IF(A1="Blad3!A6","x","y")`,
			changed: false,
		},
		{
			name:    "literal containing #REF! is not touched",
			formula: `="broken: #REF!"`,
			old:     "Blad3",
			new:     "Archief",
			want:    `="broken: #REF!"`,
			changed: false,
		},
		{
			name:    "unquoted external reference is skipped, local reference in the same formula is not",
			formula: "=[1]Blad3!A1+Blad3!A1",
			old:     "Blad3",
			new:     "Archief",
			want:    "=[1]Blad3!A1+Archief!A1",
			changed: true,
		},
		{
			name:    "quoted external reference is skipped, local reference in the same formula is not",
			formula: "=SUM('[1]Blad3'!A1,Blad3!A1)",
			old:     "Blad3",
			new:     "Archief",
			want:    "=SUM('[1]Blad3'!A1,Archief!A1)",
			changed: true,
		},
		{
			name:    "3D reference endpoint is renamed",
			formula: "=SUM(Blad3:Blad5!A1)",
			old:     "Blad3",
			new:     "Archief",
			want:    "=SUM(Archief:Blad5!A1)",
			changed: true,
		},
		{
			name:    "3D span is quoted as a whole when the new endpoint needs it",
			formula: "=SUM(Blad3:Blad5!A1:B2)",
			old:     "Blad5",
			new:     "Archief 2025",
			want:    "=SUM('Blad3:Archief 2025'!A1:B2)",
			changed: true,
		},
		{
			name:    "quoted 3D span loses its quotes when neither endpoint needs them",
			formula: "=SUM('Blad 3:Blad5'!A1)",
			old:     "Blad 3",
			new:     "Archief",
			want:    "=SUM(Archief:Blad5!A1)",
			changed: true,
		},
		{
			name:    "quoted 3D span keeps its quotes for the other endpoint",
			formula: "=SUM('Blad 3:Blad 5'!A1)",
			old:     "Blad 3",
			new:     "Archief",
			want:    "=SUM('Archief:Blad 5'!A1)",
			changed: true,
		},
		{
			name:    "structured table reference is not mistaken for a sheet reference",
			formula: "=SUM(Blad3[Column])",
			old:     "Blad3",
			new:     "Archief",
			want:    "=SUM(Blad3[Column])",
			changed: false,
		},
		{
			name:    "self-row structured reference is left alone",
			formula: "=[@Blad3]",
			old:     "Blad3",
			new:     "Archief",
			want:    "=[@Blad3]",
			changed: false,
		},
		{
			name:    "old name with an apostrophe",
			formula: "=SUM('Pete''s blad'!A1:A9)",
			old:     "Pete's blad",
			new:     "Archief",
			want:    "=SUM(Archief!A1:A9)",
			changed: true,
		},
		{
			name:    "digit-leading old name must already be quoted to match",
			formula: "='2024'!A1",
			old:     "2024",
			new:     "Archief",
			want:    "=Archief!A1",
			changed: true,
		},
		{
			name:    "accented sheet name",
			formula: "=Café!A1",
			old:     "Café",
			new:     "Archief",
			want:    "=Archief!A1",
			changed: true,
		},
		{
			name:    "no leading equals sign",
			formula: "Blad3!A1",
			old:     "Blad3",
			new:     "Archief",
			want:    "Archief!A1",
			changed: true,
		},
		{
			name:    "empty formula",
			formula: "",
			old:     "Blad3",
			new:     "Archief",
			want:    "",
			changed: false,
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			got, changed := ReplaceSheetNameInFormula(test.formula, test.old, test.new)
			if got != test.want {
				t.Errorf("formula = %q, want %q", got, test.want)
			}
			if changed != test.changed {
				t.Errorf("changed = %v, want %v", changed, test.changed)
			}
		})
	}
}

func TestQuoteSheetNameForFormula(t *testing.T) {
	tests := []struct {
		name string
		want string
	}{
		{"Blad3", "Blad3"},
		{"Archief_klanten.2024", "Archief_klanten.2024"},
		{"Archief klanten", "'Archief klanten'"},
		{"2024", "'2024'"},
		{"A1", "'A1'"},
		{"Pete's blad", "'Pete''s blad'"},
	}
	for _, test := range tests {
		if got := QuoteSheetNameForFormula(test.name); got != test.want {
			t.Errorf("QuoteSheetNameForFormula(%q) = %q, want %q", test.name, got, test.want)
		}
	}
}

func TestFormulaReferencesSheet(t *testing.T) {
	if !FormulaReferencesSheet("=SUM('Klanten actief'!A1:A9)", "Klanten actief") {
		t.Error("expected a reference to be detected")
	}
	if FormulaReferencesSheet(`=IF(A1="Klanten actief!",1,0)`, "Klanten actief") {
		t.Error("a string literal must not count as a reference")
	}
	if FormulaReferencesSheet("=[1]Blad3!A1", "Blad3") {
		t.Error("a sheet in another workbook must not count as a reference")
	}
	if !FormulaReferencesSheet("=SUM(Blad3:Blad5!A1)", "Blad3") {
		t.Error("the first endpoint of a 3D reference must count as a reference")
	}
	if !FormulaReferencesSheet("=SUM(Blad3:Blad5!A1)", "Blad5") {
		t.Error("the last endpoint of a 3D reference must count as a reference")
	}
	if FormulaReferencesSheet("=SUM(Blad3:Blad5!A1)", "Blad4") {
		t.Error("a sheet only known to lie between the 3D endpoints cannot be detected without the workbook's sheet order, so this is a documented limitation, not an assertion that it is unreachable")
	}
	if !FormulaReferencesSheet("='Blad 3:Blad 5'!A1", "Blad 3") {
		t.Error("the first endpoint of a quoted 3D reference must count as a reference")
	}
	if FormulaReferencesSheet("=SUM(Blad3[Column])", "Blad3") {
		t.Error("a structured table reference must not count as a sheet reference")
	}
	if FormulaReferencesSheet("=[@Blad3]", "Blad3") {
		t.Error("a self-row structured reference must not count as a sheet reference")
	}
}

// TestBreakReferencesToRows deletes rows 5-15 of "Data" from the point of view
// of a formula on "Data" itself and of one on "Report".
func TestBreakReferencesToRows(t *testing.T) {
	tests := []struct {
		formula string
		sheet   string
		want    string
		changed bool
	}{
		{"A6", "Data", "#REF!", true},
		{"$A$6", "Data", "#REF!", true},
		{"A4", "Data", "A4", false},
		{"A16", "Data", "A16", false},
		{"SUM(A6:A10)", "Data", "SUM(#REF!)", true},
		{"SUM(A1:A20)", "Data", "SUM(A1:A20)", false},
		{"SUM(A4:A6)", "Data", "SUM(A4:A6)", false},
		{"SUM(5:15)", "Data", "SUM(#REF!)", true},
		{"SUM(4:6)", "Data", "SUM(4:6)", false},
		{"SUM(A:A)", "Data", "SUM(A:A)", false},
		{"A1+A6+A20", "Data", "A1+#REF!+A20", true},
		{"LOG10(A6)", "Data", "LOG10(#REF!)", true},
		{"LOG10(5)", "Data", "LOG10(5)", false},
		{"ROUND(5.5,0)", "Data", "ROUND(5.5,0)", false},
		{`IF(A1="A6",A6,"")`, "Data", `IF(A1="A6",#REF!,"")`, true},
		{"A6", "Report", "A6", false},
		{"Data!A6", "Report", "#REF!", true},
		{"'Data'!A6:B7", "Report", "#REF!", true},
		{"data!A6+Data!A20", "Report", "#REF!+Data!A20", true},
		{"Report!A6", "Report", "Report!A6", false},
		{"[1]Data!A6", "Report", "[1]Data!A6", false},
		{"Data!#REF!", "Report", "Data!#REF!", false},
		{"Data!MyName", "Report", "Data!MyName", false},
		// 4-letter columns can't be real (max is XFD), so this parses as a
		// name, not a reference, and must be left alone.
		{"AAAA6", "Data", "AAAA6", false},
		// A row number 8 digits long is already beyond Excel's 1048576-row
		// limit, so this can't be a real row reference either.
		{"A10485766", "Data", "A10485766", false},
		{"Tab6[Column]", "Data", "Tab6[Column]", false},
		{"[@A6]", "Data", "[@A6]", false},
		{"Data:Report!A6", "Report", "Data:Report!A6", false},
	}
	for _, test := range tests {
		got, changed := BreakReferencesToRows(test.formula, test.sheet, "Data", 5, 15)
		if got != test.want || changed != test.changed {
			t.Errorf("BreakReferencesToRows(%q on %s) = %q, %v; want %q, %v", test.formula, test.sheet, got, changed, test.want, test.changed)
		}
	}
}

func TestShiftFormulaReferences(t *testing.T) {
	tests := []struct {
		formula    string
		dCol, dRow int
		want       string
	}{
		{"A1*2", 0, 3, "A4*2"},
		{"A1*2", 2, 0, "C1*2"},
		{"$A$1+A$1+$A1", 1, 1, "$A$1+B$1+$A2"},
		{"SUM(A1:B2)", 1, 1, "SUM(B2:C3)"},
		{"Data!A1+'My sheet'!B2", 0, 1, "Data!A2+'My sheet'!B3"},
		// A sheet-qualified relative reference must shift its cell part the
		// same as a bare one: filling =Data!A1 down a row gives =Data!A2 in
		// real Excel. excelize's own shared-formula shifter gets this wrong
		// (it never touches the cell part of a sheet-qualified operand), so
		// this is pinned down explicitly rather than trusted by inspection.
		{"Data!A1", 0, 1, "Data!A2"},
		{"'Bron gegevens'!A1", 0, 1, "'Bron gegevens'!A2"},
		{"Data!$A$1", 1, 1, "Data!$A$1"},
		{"Data!A$1", 1, 1, "Data!B$1"},
		{"Data!$A1", 1, 1, "Data!$A2"},
		{"Data!A1:B2", 1, 1, "Data!B2:C3"},
		{`IF(A1="A1",A1,"")`, 0, 1, `IF(A2="A1",A2,"")`},
		{"LOG10(A1)+PI()", 0, 1, "LOG10(A2)+PI()"},
		{"SUM(A:A)+SUM(1:1)", 1, 1, "SUM(B:B)+SUM(2:2)"},
		{"A1+MyName+1.5", 0, 1, "A2+MyName+1.5"},
		{"A1", 0, -1, "#REF!"},
		{"A1", -1, 0, "#REF!"},
		// A table column can be named like a cell (e.g. a column literally
		// called "A1"); [@A1] must never be read as a shiftable reference.
		{"[@A1]", 0, 1, "[@A1]"},
		{"[@Column]", 1, 1, "[@Column]"},
		{"Table1[Column]", 1, 1, "Table1[Column]"},
		// A short table name that happens to look like a cell address (e.g.
		// "Tab1") must not be shifted, let alone reformatted to uppercase.
		{"Tab1[Column]", 0, 1, "Tab1[Column]"},
		// A shared formula group can reach into another workbook; Excel
		// shifts those relative references exactly like local ones.
		{"[1]Data!A1", 0, 1, "[1]Data!A2"},
		{"'[1]My data'!A1", 0, 1, "'[1]My data'!A2"},
		// Already out of Excel's limits before any shift is applied.
		{"AAAA1", 0, 1, "AAAA1"},
		{"A10485766", 0, 0, "A10485766"},
	}
	for _, test := range tests {
		if got := ShiftFormulaReferences(test.formula, test.dCol, test.dRow); got != test.want {
			t.Errorf("ShiftFormulaReferences(%q, %d, %d) = %q, want %q", test.formula, test.dCol, test.dRow, got, test.want)
		}
	}
}

// TestRewriteFormulaReferencesRoundTrip is the strongest guard the walker
// has: with a visitor that never substitutes anything, the output must equal
// the input byte for byte, for every formula, however strange. If this ever
// fails, the walker has started dropping or reordering text it does not
// understand, which is exactly the kind of silent corruption this file must
// never produce.
func TestRewriteFormulaReferencesRoundTrip(t *testing.T) {
	corpus := []string{
		// String literals containing ref-like text.
		`="A6"`,
		`=IF(A1="Sheet1!A6","x","y")`,
		`=""""`,
		`="broken: #REF!"`,
		`=CONCAT("a""Blad3!b",Blad3!C1)`,
		// Hostile sheet names.
		"='Pete''s blad'!A6",
		"='Sheet with spaces'!A6",
		"='Sheet!bang'!A6",
		"='2024'!A6",
		"='A1'!A6",
		"='REF'!A6",
		"=Café!A6",
		"='" + strings.Repeat("x", 31) + "'!A6",
		// External workbook references.
		"=[1]Sheet1!A6",
		"='[1]Some sheet'!A6",
		"='[1]Some sheet'!A6:B7",
		// 3D references.
		"=Sheet1:Sheet3!A1",
		"='Sheet 1:Sheet 3'!A1",
		// Structured/table references.
		"=Table1[#All]",
		"=Table1[Column]",
		"=[@Column]",
		"=[@A1]",
		"=Table1[[#Headers],[Column1]:[Column2]]",
		"=Tab1[Column]",
		// Defined names and functions that look like references.
		"=LOG10(A6)",
		"=ATAN2(1,2)",
		"=A1(",
		// Absolute and mixed references.
		"=$A$6+A$6+$A6",
		// Whole row/column references.
		"=SUM(5:15)",
		"=SUM(A:A)",
		"=SUM($A:$A)",
		// Ranges overlapping a deleted block in every possible way.
		"=SUM(A1:A20)",
		"=SUM(A4:A6)",
		"=SUM(A6:A10)",
		// Beyond Excel's limits.
		"=A1048577",
		"=XFE1",
		"=AAAA1",
		// Error literals.
		"=#REF!",
		"=#N/A",
		"=#DIV/0!",
		"=#NAME?",
		"=Sheet1!#REF!",
		// No leading "=", empty, whitespace, newlines.
		"A1+B1",
		"",
		"   ",
		"=A1+\nB1",
		// Array formulas, intersection and union.
		"{=SUM(A1:A6*B1:B6)}",
		"=SUM(A1:A6 B1:B6)",
		"=SUM(A1,A6)",
		// Range with a function as its second half.
		"=SUM(Blad3!A1:INDEX(Blad3!A:A,3))",
		// Deep nesting.
		strings.Repeat("(", 500) + "A1" + strings.Repeat(")", 500),
	}
	for _, formula := range corpus {
		t.Run(formula, func(t *testing.T) {
			got, changed := rewriteFormulaReferences(formula, func(formulaReference) (string, bool) {
				return "", false
			})
			if got != formula {
				t.Errorf("rewriteFormulaReferences(%q) = %q, want the input unchanged", formula, got)
			}
			if changed {
				t.Errorf("rewriteFormulaReferences(%q) reported changed=true with a no-op visitor", formula)
			}
		})
	}
}

// TestRewriteFormulaReferencesLargeFormula guards against accidental
// quadratic behaviour or unbounded lookahead: this must complete quickly and
// still round-trip cleanly.
func TestRewriteFormulaReferencesLargeFormula(t *testing.T) {
	var b strings.Builder
	b.WriteString("=SUM(")
	for i := 1; i <= 20000; i++ {
		if i > 1 {
			b.WriteString(",")
		}
		b.WriteString("Sheet1!A" + strconv.Itoa(i))
	}
	b.WriteString(")")
	formula := b.String()

	got, changed := rewriteFormulaReferences(formula, func(formulaReference) (string, bool) {
		return "", false
	})
	if got != formula || changed {
		t.Errorf("a large formula was not round-tripped unchanged")
	}

	renamed, ok := ReplaceSheetNameInFormula(formula, "Sheet1", "Archief")
	if !ok || strings.Contains(renamed, "Sheet1!") || !strings.Contains(renamed, "Archief!A1,") {
		t.Errorf("a large formula was not renamed correctly")
	}
}

// TestReadErrorLiteralVariants locks in exactly which trailing punctuation
// belongs to each error literal, since readErrorLiteral has to draw that
// line without a table of known error names.
func TestReadErrorLiteralVariants(t *testing.T) {
	tests := []struct {
		formula string
		want    string
	}{
		{"=#REF!+A1", "=#REF!+A1"},
		{"=#N/A+A1", "=#N/A+A1"},
		{"=#DIV/0!+A1", "=#DIV/0!+A1"},
		{"=#NAME?+A1", "=#NAME?+A1"},
		{"=#NULL!+A1", "=#NULL!+A1"},
		{"=#NUM!+A1", "=#NUM!+A1"},
		{"=#VALUE!+A1", "=#VALUE!+A1"},
		{"=#GETTING_DATA", "=#GETTING_DATA"},
	}
	for _, test := range tests {
		got, changed := rewriteFormulaReferences(test.formula, func(formulaReference) (string, bool) {
			return "", false
		})
		if got != test.want || changed {
			t.Errorf("rewriteFormulaReferences(%q) = %q, changed=%v; want %q, changed=false", test.formula, got, changed, test.want)
		}
	}
}
