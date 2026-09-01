package excel

import "testing"

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
}
