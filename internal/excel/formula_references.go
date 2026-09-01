package excel

import (
	"regexp"
	"strconv"
	"strings"
	"unicode"

	"github.com/xuri/excelize/v2"
)

// unquotedSheetNamePattern matches sheet names that may appear in a formula
// without surrounding single quotes.
var unquotedSheetNamePattern = regexp.MustCompile(`^[\p{L}_][\p{L}\p{N}_.]*$`)

// cellReferencePattern matches names that would be ambiguous with a cell
// reference (e.g. "A1"), which always need quoting when used as a sheet name.
var cellReferencePattern = regexp.MustCompile(`^[A-Za-z]{1,3}[0-9]{1,7}$`)

// referencePartPattern splits one side of a cell or range reference into its
// column and row, each with an optional "$" anchor. Either side may be empty
// so that whole-column ("A:A") and whole-row ("5:10") references parse too.
var referencePartPattern = regexp.MustCompile(`^(\$?)([A-Za-z]{0,3})(\$?)([0-9]{0,7})$`)

// QuoteSheetNameForFormula renders a sheet name as it must appear in front of
// the "!" in a formula, adding single quotes only when they are required.
func QuoteSheetNameForFormula(name string) string {
	if unquotedSheetNamePattern.MatchString(name) && !cellReferencePattern.MatchString(name) {
		return name
	}
	return "'" + strings.ReplaceAll(name, "'", "''") + "'"
}

// formulaReference is one reference found inside a formula: an optional sheet
// qualifier plus the cell, range or name it points at.
type formulaReference struct {
	// Qualifier is the sheet part exactly as written, quotes included, or ""
	// when the reference is unqualified.
	Qualifier string
	// Sheet is the unquoted sheet name, "" when the reference is unqualified.
	Sheet string
	// External marks a reference into another workbook, such as [1]Sheet1!A1,
	// which no sheet operation on this workbook may touch.
	External bool
	// Ref is the part after the "!", or the whole token when unqualified.
	Ref string
}

// isReferenceRune reports whether r may appear in an unquoted sheet name, a
// cell reference or a function/defined name inside a formula.
func isReferenceRune(r rune) bool {
	return r == '_' || r == '.' || r == '$' || unicode.IsLetter(r) || unicode.IsDigit(r)
}

// rewriteFormulaReferences walks formula and offers every reference to visit,
// which returns replacement text and true to substitute it. String literals
// and error literals are skipped and everything else is copied verbatim, so a
// formula in which nothing is replaced comes back byte for byte identical.
//
// This is the one place that understands formula syntax; sheet renames, row
// deletions and shared formula expansion are all visitors on top of it.
func rewriteFormulaReferences(formula string, visit func(ref formulaReference) (string, bool)) (string, bool) {
	var out strings.Builder
	changed := false
	runes := []rune(formula)
	// prev is the last significant rune copied; "]" means the token we are
	// about to read belongs to an external workbook reference.
	var prev rune

	// readReference consumes a cell, range or name token starting at i. A
	// ":" joins two halves into a range unless the second half opens a
	// function call, as in A1:INDEX(...).
	readReference := func(i int) int {
		for i < len(runes) && isReferenceRune(runes[i]) {
			i++
		}
		if i+1 < len(runes) && runes[i] == ':' && isReferenceRune(runes[i+1]) {
			end := i + 1
			for end < len(runes) && isReferenceRune(runes[end]) {
				end++
			}
			if end >= len(runes) || runes[end] != '(' {
				return end
			}
		}
		return i
	}
	emit := func(ref formulaReference, original string) {
		if replacement, ok := visit(ref); ok {
			out.WriteString(replacement)
			changed = true
		} else {
			out.WriteString(original)
		}
	}
	// emitQualified handles the part after a sheet qualifier's "!", which may
	// be a cell, a range, a name or an error literal such as Sheet1!#REF!.
	emitQualified := func(ref formulaReference, start int, i int) int {
		refStart := i
		if i < len(runes) && runes[i] == '#' {
			i = readErrorLiteral(runes, i)
		} else {
			i = readReference(i)
		}
		ref.Ref = string(runes[refStart:i])
		emit(ref, string(runes[start:i]))
		return i
	}

	for i := 0; i < len(runes); {
		switch r := runes[i]; {
		case r == '"':
			// String literal: copy verbatim, "" is an escaped quote.
			end := i + 1
			for end < len(runes) {
				if runes[end] == '"' {
					if end+1 < len(runes) && runes[end+1] == '"' {
						end += 2
						continue
					}
					end++
					break
				}
				end++
			}
			out.WriteString(string(runes[i:end]))
			i, prev = end, '"'

		case r == '#':
			end := readErrorLiteral(runes, i)
			out.WriteString(string(runes[i:end]))
			i, prev = end, runes[end-1]

		case r == '[':
			// Balanced bracket content: a structured table reference such as
			// Table1[Column], [@Column] or [#Headers], or an external
			// workbook index such as the "[1]" in [1]Sheet1!A6. A table
			// column can itself be named like a cell (e.g. [@A1] refers to a
			// column literally named "A1"), so its contents must never be
			// run through the reference tokenizer below. depth handles the
			// nested form Table1[[#Headers],[Column1]:[Column2]].
			start := i
			depth := 0
			for i < len(runes) {
				if runes[i] == '[' {
					depth++
				} else if runes[i] == ']' {
					depth--
					if depth == 0 {
						i++
						break
					}
				}
				i++
			}
			out.WriteString(string(runes[start:i]))
			prev = ']'

		case r == '\'':
			// Quoted sheet name, '' is an escaped quote.
			start := i
			var name strings.Builder
			i++
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
			if closed && i < len(runes) && runes[i] == '!' {
				ref := formulaReference{
					Qualifier: string(runes[start:i]),
					Sheet:     name.String(),
					External:  strings.HasPrefix(name.String(), "["),
				}
				i = emitQualified(ref, start, i+1)
			} else {
				out.WriteString(string(runes[start:i]))
			}
			prev = '\''

		case isReferenceRune(r):
			start := i
			i = readReference(i)
			token := string(runes[start:i])
			switch {
			case i < len(runes) && runes[i] == '!':
				ref := formulaReference{Qualifier: token, Sheet: token, External: prev == ']'}
				i = emitQualified(ref, start, i+1)
			case i < len(runes) && runes[i] == '(':
				// A function call, never a reference.
				out.WriteString(token)
			case i < len(runes) && runes[i] == '[':
				// A table name in front of a structured reference, as in
				// Table1[Column]. Table names are validated against the same
				// "must not look like a cell reference" rule as sheet names,
				// but that rule is not enforced here, so a short letters-then-
				// digits table name (e.g. "Tab1") would otherwise be
				// misread as cell TAB1 and shifted/broken like one. Never a
				// reference.
				out.WriteString(token)
			default:
				emit(formulaReference{Ref: token}, token)
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

// readErrorLiteral returns the index just past the error literal that starts
// at i, e.g. #REF!, #N/A or #DIV/0!.
func readErrorLiteral(runes []rune, i int) int {
	i++
	for i < len(runes) && (runes[i] == '/' || runes[i] == '_' || unicode.IsLetter(runes[i]) || unicode.IsDigit(runes[i])) {
		i++
	}
	if i < len(runes) && (runes[i] == '!' || runes[i] == '?') {
		i++
	}
	return i
}

// ReplaceSheetNameInFormula rewrites every reference to oldName in an Excel
// formula so it points at newName, and reports whether anything changed.
//
// Excelize's SetSheetName deliberately leaves formulas alone, so without this
// a rename silently turns every cross-sheet formula into a broken reference.
//
// A 3D reference such as Sheet1:Sheet3!A1 is left untouched even when oldName
// is one of its endpoints: ref.Sheet holds the whole "Sheet1:Sheet3" span, so
// it never equals oldName and this function safely no-ops on it rather than
// risk misquoting the pair. Known incompleteness, not corruption: the rename
// silently fails to reach a 3D reference; see FormulaReferencesSheet for the
// case that matters more, refusing to delete a sheet that is still in use.
func ReplaceSheetNameInFormula(formula string, oldName string, newName string) (string, bool) {
	replacement := QuoteSheetNameForFormula(newName) + "!"
	return rewriteFormulaReferences(formula, func(ref formulaReference) (string, bool) {
		if ref.External || !strings.EqualFold(ref.Sheet, oldName) {
			return "", false
		}
		return replacement + ref.Ref, true
	})
}

// FormulaReferencesSheet reports whether the formula refers to the given sheet.
//
// A 3D reference such as Sheet1:Sheet3!A1 covers every sheet between its two
// endpoints in tab order, but this function has no access to that order, so
// it can only recognise sheetName as one of the two named endpoints, not as
// a sheet lying between them. That is the best this signature can do; a
// sheet named exactly at either end is still caught, which is what stops a
// delete-sheet call from silently breaking that reference.
func FormulaReferencesSheet(formula string, sheetName string) bool {
	found := false
	rewriteFormulaReferences(formula, func(ref formulaReference) (string, bool) {
		if ref.External {
			return "", false
		}
		first, last, _ := strings.Cut(ref.Sheet, ":")
		if strings.EqualFold(first, sheetName) || (last != "" && strings.EqualFold(last, sheetName)) {
			found = true
		}
		return "", false
	})
	return found
}

// BreakReferencesToRows replaces every reference that lies entirely inside
// rows startRow..endRow of sheetName with #REF!, which is what Excel does when
// those rows are deleted. formulaSheet is the sheet holding the formula, and
// so the target of its unqualified references. References that only overlap
// the rows are left alone: they shrink rather than break.
//
// The sheet qualifier is dropped even though Excel itself writes Sheet1!#REF!:
// excelize's formula parser cannot read that form and then leaves every other
// reference in the formula unadjusted when the rows are removed. A bare #REF!
// evaluates identically and keeps the rest of the formula correct.
//
// A 3D reference such as Sheet1:Sheet3!A6 is left untouched: ref.Sheet holds
// the whole "Sheet1:Sheet3" span, which never equals a plain sheetName, so
// this safely no-ops on it instead of guessing whether sheetName falls
// inside the range.
func BreakReferencesToRows(formula string, formulaSheet string, sheetName string, startRow int, endRow int) (string, bool) {
	return rewriteFormulaReferences(formula, func(ref formulaReference) (string, bool) {
		target := ref.Sheet
		if ref.Qualifier == "" {
			target = formulaSheet
		}
		if ref.External || !strings.EqualFold(target, sheetName) {
			return "", false
		}
		first, last, ok := referenceRows(ref.Ref)
		if !ok || first < startRow || last > endRow {
			return "", false
		}
		return "#REF!", true
	})
}

// referenceRows returns the first and last row a cell, cell range or row range
// covers. Whole-column references and anything that is not a reference report
// false.
func referenceRows(ref string) (int, int, bool) {
	first, last := 0, 0
	parts := strings.Split(ref, ":")
	if len(parts) > 2 {
		return 0, 0, false
	}
	for i, part := range parts {
		match := referencePartPattern.FindStringSubmatch(part)
		if match == nil || match[4] == "" {
			return 0, 0, false
		}
		// A lone "5" is a number, not row 5; only a range may omit columns.
		if match[2] == "" && len(parts) == 1 {
			return 0, 0, false
		}
		row, _ := strconv.Atoi(match[4])
		if i == 0 {
			first, last = row, row
		} else {
			first, last = min(first, row), max(last, row)
		}
	}
	return first, last, true
}

// ShiftFormulaReferences moves every relative reference in formula by dCol
// columns and dRow rows, honouring "$" anchors. This is how a shared formula
// group derives each member's formula from the master's. A reference that
// would move off the sheet becomes #REF!.
func ShiftFormulaReferences(formula string, dCol int, dRow int) string {
	shifted, _ := rewriteFormulaReferences(formula, func(ref formulaReference) (string, bool) {
		parts := strings.Split(ref.Ref, ":")
		if len(parts) > 2 {
			return "", false
		}
		for i, part := range parts {
			match := referencePartPattern.FindStringSubmatch(part)
			if match == nil || (match[2] == "" && match[4] == "") {
				return "", false
			}
			// A lone column or row is only a reference inside a range.
			if len(parts) == 1 && (match[2] == "" || match[4] == "") {
				return "", false
			}
			column, row := match[2], match[4]
			if column != "" && match[1] == "" {
				number, _ := excelize.ColumnNameToNumber(column)
				if number += dCol; number < 1 || number > excelize.MaxColumns {
					return "#REF!", true
				}
				column, _ = excelize.ColumnNumberToName(number)
			}
			if row != "" && match[3] == "" {
				number, _ := strconv.Atoi(row)
				if number += dRow; number < 1 || number > excelize.TotalRows {
					return "#REF!", true
				}
				row = strconv.Itoa(number)
			}
			parts[i] = match[1] + column + match[3] + row
		}
		return ref.Qualifier + qualifierSeparator(ref) + strings.Join(parts, ":"), true
	})
	return shifted
}

func qualifierSeparator(ref formulaReference) string {
	if ref.Qualifier == "" {
		return ""
	}
	return "!"
}
