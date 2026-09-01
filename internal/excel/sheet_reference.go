package excel

import (
	"regexp"
	"strings"
	"unicode"
)

// unquotedSheetNamePattern matches sheet names that may appear in a formula
// without surrounding single quotes.
var unquotedSheetNamePattern = regexp.MustCompile(`^[\p{L}_][\p{L}\p{N}_.]*$`)

// cellReferencePattern matches names that would be ambiguous with a cell
// reference (e.g. "A1"), which always need quoting when used as a sheet name.
var cellReferencePattern = regexp.MustCompile(`^[A-Za-z]{1,3}[0-9]{1,7}$`)

// QuoteSheetNameForFormula renders a sheet name as it must appear in front of
// the "!" in a formula, adding single quotes only when they are required.
func QuoteSheetNameForFormula(name string) string {
	if unquotedSheetNamePattern.MatchString(name) && !cellReferencePattern.MatchString(name) {
		return name
	}
	return "'" + strings.ReplaceAll(name, "'", "''") + "'"
}

// isSheetNameRune reports whether r may appear in an unquoted sheet name or in
// a function/defined name token inside a formula.
func isSheetNameRune(r rune) bool {
	return r == '_' || r == '.' || unicode.IsLetter(r) || unicode.IsDigit(r)
}

// ReplaceSheetNameInFormula rewrites every reference to oldName in an Excel
// formula so it points at newName, and reports whether anything changed.
//
// Excelize's SetSheetName deliberately leaves formulas alone, so without this
// a rename silently turns every cross-sheet formula into a broken reference.
// String literals and references into other workbooks ("[1]Sheet1!A1") are left
// untouched.
func ReplaceSheetNameInFormula(formula string, oldName string, newName string) (string, bool) {
	if formula == "" {
		return formula, false
	}
	replacement := QuoteSheetNameForFormula(newName)

	var out strings.Builder
	changed := false
	runes := []rune(formula)
	// prev is the last significant rune emitted; "]" means the token we are
	// about to read belongs to an external workbook reference.
	var prev rune

	for i := 0; i < len(runes); {
		switch r := runes[i]; {
		case r == '"':
			// String literal: copy verbatim, "" is an escaped quote.
			out.WriteRune(r)
			i++
			for i < len(runes) {
				out.WriteRune(runes[i])
				if runes[i] == '"' {
					if i+1 < len(runes) && runes[i+1] == '"' {
						out.WriteRune(runes[i+1])
						i += 2
						continue
					}
					i++
					break
				}
				i++
			}
			prev = '"'

		case r == '\'':
			// Quoted sheet name, '' is an escaped quote.
			start := i
			i++
			var name strings.Builder
			closed := false
			for i < len(runes) {
				if runes[i] == '\'' {
					if i+1 < len(runes) && runes[i+1] == '\'' {
						name.WriteRune('\'')
						i += 2
						continue
					}
					i++
					closed = true
					break
				}
				name.WriteRune(runes[i])
				i++
			}
			followedByBang := closed && i < len(runes) && runes[i] == '!'
			if followedByBang && strings.EqualFold(name.String(), oldName) {
				out.WriteString(replacement)
				changed = true
			} else {
				out.WriteString(string(runes[start:i]))
			}
			prev = '\''

		case isSheetNameRune(r):
			start := i
			for i < len(runes) && isSheetNameRune(runes[i]) {
				i++
			}
			token := string(runes[start:i])
			followedByBang := i < len(runes) && runes[i] == '!'
			// prev == ']' means "[1]Sheet1!A1": another workbook's sheet.
			if followedByBang && prev != ']' && strings.EqualFold(token, oldName) {
				out.WriteString(replacement)
				changed = true
			} else {
				out.WriteString(token)
			}
			prev = runes[i-1]

		default:
			out.WriteRune(r)
			if !unicode.IsSpace(r) {
				prev = r
			}
			i++
		}
	}
	return out.String(), changed
}

// FormulaReferencesSheet reports whether the formula refers to the given sheet.
func FormulaReferencesSheet(formula string, sheetName string) bool {
	// Rewriting to a name that cannot collide tells us whether a reference
	// exists without duplicating the tokenizer.
	_, found := ReplaceSheetNameInFormula(formula, sheetName, sheetName+"\x00")
	return found
}
