package tools

import (
	"context"
	"fmt"
	"html"
	"strings"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	excel "github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ExcelRenameSheetArguments struct {
	FileAbsolutePath string `zog:"fileAbsolutePath"`
	SheetName        string `zog:"sheetName"`
	NewName          string `zog:"newName"`
}

var excelRenameSheetArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"newName":          z.String().Required(),
})

func AddExcelRenameSheetTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_rename_sheet",
		mcp.WithDescription("Rename a sheet in the Excel file. Cell values, formatting, merged cells, column widths and row heights are preserved. Formulas, defined names, data validation lists, chart series, table column formulas and pivot table sources referring to the sheet are updated; conditional formatting rules referring to it are reported as a warning."),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Current name of the sheet to rename"),
		),
		mcp.WithString("newName",
			mcp.Required(),
			mcp.Description("New name for the sheet. Maximum 31 characters, and it cannot contain : \\ / ? * [ ]"),
		),
	), handleRenameSheet)
}

func handleRenameSheet(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelRenameSheetArguments{}
	if issues := excelRenameSheetArgumentsSchema.Parse(request.Params.Arguments, &args); len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return renameSheet(args.FileAbsolutePath, args.SheetName, args.NewName)
}

func renameSheet(fileAbsolutePath string, sheetName string, newName string) (*mcp.CallToolResult, error) {
	if err := validateSheetName(newName); err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}

	workbook, release, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer release()

	warnings, err := workbook.RenameSheet(sheetName, newName)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	if err := workbook.Save(); err != nil {
		return nil, err
	}

	result := "# Notice\n"
	result += fmt.Sprintf("backend: %s\n", workbook.GetBackendName())
	result += fmt.Sprintf("Sheet [%s] renamed to [%s].\n", html.EscapeString(sheetName), html.EscapeString(newName))
	result += formatSheetOrder(workbook)
	result += formatWarnings(warnings)
	return mcp.NewToolResultText(result), nil
}

// validateSheetName applies the workbook-format rules that Excel enforces on
// sheet names, so the caller gets a clear message instead of a backend error.
func validateSheetName(name string) error {
	if strings.TrimSpace(name) == "" {
		return fmt.Errorf("sheet name must not be empty")
	}
	if len([]rune(name)) > 31 {
		return fmt.Errorf("sheet name [%s] is %d characters long, but Excel allows at most 31", name, len([]rune(name)))
	}
	if strings.ContainsAny(name, `:\/?*[]`) {
		return fmt.Errorf(`sheet name [%s] contains a character that Excel does not allow (: \ / ? * [ ])`, name)
	}
	if strings.HasPrefix(name, "'") || strings.HasSuffix(name, "'") {
		return fmt.Errorf("sheet name [%s] must not start or end with a single quote", name)
	}
	return nil
}

// formatSheetOrder renders the resulting sheet order so the caller can verify
// the workbook state without a follow-up describe call.
func formatSheetOrder(workbook excel.Excel) string {
	names, err := workbook.SheetNames()
	if err != nil {
		return ""
	}
	result := "Sheets (in order):\n"
	for i, name := range names {
		result += fmt.Sprintf("%d. %s\n", i, html.EscapeString(name))
	}
	return result
}

func formatWarnings(warnings []string) string {
	if len(warnings) == 0 {
		return ""
	}
	result := "\n# Warnings\n"
	for _, warning := range warnings {
		result += fmt.Sprintf("- %s\n", html.EscapeString(warning))
	}
	return result
}
