package tools

import (
	"testing"
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
