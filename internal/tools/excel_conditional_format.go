package tools

import (
	"context"
	"fmt"
	"html"
	"strings"

	z "github.com/Oudwins/zog"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"github.com/negokaz/excel-mcp-server/internal/excel"
	imcp "github.com/negokaz/excel-mcp-server/internal/mcp"
)

type ExcelConditionalFormatArguments struct {
	FileAbsolutePath string                         `zog:"fileAbsolutePath"`
	SheetName        string                         `zog:"sheetName"`
	Range            string                         `zog:"range"`
	Operation        string                         `zog:"operation"`
	Rules            []*excel.ConditionalFormatRule `zog:"rules"`
}

var conditionalFormatRuleTypes = []string{
	"cellIs", "expression", "containsText", "timePeriod",
	"top10", "aboveAverage", "duplicate", "unique",
	"2_color_scale", "3_color_scale", "dataBar",
}

var conditionalFormatOperators = []string{
	"lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual",
	"equal", "notEqual", "between", "notBetween",
	"containsText", "notContains", "beginsWith", "endsWith",
}

var excelConditionalFormatArgumentsSchema = z.Struct(z.Shape{
	"fileAbsolutePath": z.String().Test(AbsolutePathTest()).Required(),
	"sheetName":        z.String().Required(),
	"range":            z.String().Required(),
	"operation":        z.String().OneOf([]string{"set", "clear"}).Default("set"),
	"rules": z.Slice(z.Ptr(z.Struct(z.Shape{
		"type":       z.String().OneOf(conditionalFormatRuleTypes).Required(),
		"operator":   z.String().OneOf(conditionalFormatOperators),
		"formulas":   z.Slice(z.String()).Default([]string{}),
		"text":       z.String(),
		"stopIfTrue": z.Bool().Default(false),
		"font": z.Ptr(z.Struct(z.Shape{
			"bold":   z.Ptr(z.Bool()),
			"italic": z.Ptr(z.Bool()),
			"strike": z.Ptr(z.Bool()),
			"color":  z.Ptr(z.String().Match(colorPattern)),
		})),
		"fill": z.Ptr(z.Struct(z.Shape{
			"type":    z.StringLike[excel.FillType]().OneOf(excel.FillTypeValues()).Default(excel.FillTypePattern),
			"pattern": z.StringLike[excel.FillPattern]().OneOf(excel.FillPatternValues()).Default(excel.FillPatternSolid),
			"color":   z.Slice(z.String().Match(colorPattern)).Default([]string{}),
		})),
	}))).Default([]*excel.ConditionalFormatRule{}),
})

func AddExcelConditionalFormatTool(server *server.MCPServer) {
	server.AddTool(mcp.NewTool("excel_conditional_format",
		mcp.WithDescription(
			"Set or clear conditional formatting rules on a range, such as colouring a cell by how its date compares to TODAY(). "+
				"Rules are evaluated in the order given. Use excel_read_sheet with showStyle to see the rules a sheet already has. "+
				"Setting rules on a range replaces the rules previously written to that same range."),
		mcp.WithString("fileAbsolutePath",
			mcp.Required(),
			mcp.Description("Absolute path to the Excel file"),
		),
		mcp.WithString("sheetName",
			mcp.Required(),
			mcp.Description("Sheet name in the Excel file"),
		),
		mcp.WithString("range",
			mcp.Required(),
			mcp.Description("Range the rules apply to (e.g., \"A1:A11\")"),
		),
		mcp.WithString("operation",
			mcp.Description("\"set\" to write the given rules, \"clear\" to remove the conditional formatting on the range. Defaults to \"set\"."),
			mcp.Enum("set", "clear"),
		),
		mcp.WithArray("rules",
			mcp.Description("Rules to apply, in evaluation order. Required when operation is \"set\"."),
			mcp.Items(map[string]any{
				"type": "object",
				"properties": map[string]any{
					"type": map[string]any{
						"type": "string",
						"enum": conditionalFormatRuleTypes,
						"description": "Rule type. \"cellIs\" compares the cell against formulas using operator; " +
							"\"expression\" applies when its single formula is true.",
					},
					"operator": map[string]any{
						"type":        "string",
						"enum":        conditionalFormatOperators,
						"description": "Comparison for cellIs and text rules. Not used by expression rules.",
					},
					"formulas": map[string]any{
						"type":  "array",
						"items": map[string]any{"type": "string"},
						"description": "Values or formulas the rule compares against, without a leading \"=\". " +
							"One entry for most operators, two for between/notBetween, and exactly one for an expression rule " +
							"(e.g. \"$C1<TODAY()\").",
					},
					"text": map[string]any{
						"type":        "string",
						"description": "The text a containsText rule looks for.",
					},
					"stopIfTrue": map[string]any{
						"type":        "boolean",
						"description": "Stop evaluating later rules on this range once this one matches.",
					},
					"font": map[string]any{
						"type":        "object",
						"description": "Font applied when the rule matches.",
						"properties": map[string]any{
							"bold":   map[string]any{"type": "boolean"},
							"italic": map[string]any{"type": "boolean"},
							"strike": map[string]any{"type": "boolean"},
							"color":  map[string]any{"type": "string", "pattern": colorPattern.String()},
						},
					},
					"fill": map[string]any{
						"type":        "object",
						"description": "Background fill applied when the rule matches.",
						"properties": map[string]any{
							"type":    map[string]any{"type": "string", "enum": []string{"pattern", "gradient"}},
							"pattern": map[string]any{"type": "string", "enum": excel.FillPatternValues()},
							"color": map[string]any{
								"type":  "array",
								"items": map[string]any{"type": "string", "pattern": colorPattern.String()},
							},
						},
					},
				},
				"required": []string{"type"},
			}),
		),
	), handleConditionalFormat)
}

func handleConditionalFormat(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args := ExcelConditionalFormatArguments{}
	issues := excelConditionalFormatArgumentsSchema.Parse(request.Params.Arguments, &args)
	if len(issues) != 0 {
		return imcp.NewToolResultZogIssueMap(issues), nil
	}
	return conditionalFormat(args.FileAbsolutePath, args.SheetName, args.Range, args.Operation, args.Rules)
}

func conditionalFormat(fileAbsolutePath, sheetName, rangeStr, operation string, rules []*excel.ConditionalFormatRule) (*mcp.CallToolResult, error) {
	if operation == "set" && len(rules) == 0 {
		return imcp.NewToolResultInvalidArgumentError("operation \"set\" needs at least one rule; use operation \"clear\" to remove the conditional formatting on a range"), nil
	}
	if _, _, _, _, err := excel.ParseRange(rangeStr); err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}

	workbook, closeFn, err := excel.OpenFile(fileAbsolutePath)
	if err != nil {
		return nil, err
	}
	defer closeFn()

	worksheet, err := workbook.FindSheet(sheetName)
	if err != nil {
		return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
	}
	defer worksheet.Release()

	switch operation {
	case "clear":
		if err := worksheet.ClearConditionalFormat(rangeStr); err != nil {
			return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
		}
	default:
		applied := make([]excel.ConditionalFormatRule, 0, len(rules))
		for i, rule := range rules {
			if rule == nil {
				return imcp.NewToolResultInvalidArgumentError(fmt.Sprintf("rule %d is null", i+1)), nil
			}
			applied = append(applied, *rule)
		}
		if err := worksheet.SetConditionalFormat(rangeStr, applied); err != nil {
			return imcp.NewToolResultInvalidArgumentError(err.Error()), nil
		}
	}

	if err := workbook.Save(); err != nil {
		return nil, err
	}

	// Report the rules as they now read back from the file, rather than as
	// they were asked for: that is what Excel will act on, and it surfaces
	// anything the backend altered on the way in.
	stored, err := worksheet.GetConditionalFormats()
	if err != nil {
		return nil, err
	}

	result := "<h2>Conditional Formatting</h2>\n"
	if operation == "clear" {
		result += fmt.Sprintf("<p>Cleared the conditional formatting on %s in sheet %s</p>\n", html.EscapeString(rangeStr), html.EscapeString(sheetName))
	} else {
		result += fmt.Sprintf("<p>Applied %d rule(s) to %s in sheet %s</p>\n", len(rules), html.EscapeString(rangeStr), html.EscapeString(sheetName))
	}
	result += "<h2>Rules now on this sheet</h2>\n<ul>\n"
	if len(stored) == 0 {
		result += "<li>none</li>\n"
	}
	for _, rule := range stored {
		parts := []string{html.EscapeString(rule.Range), "type=" + html.EscapeString(rule.Type)}
		if rule.Operator != "" {
			parts = append(parts, "operator="+html.EscapeString(rule.Operator))
		}
		if len(rule.Formulas) > 0 {
			parts = append(parts, "formula=["+html.EscapeString(strings.Join(rule.Formulas, ", "))+"]")
		}
		if rule.Fill != nil && len(rule.Fill.Color) > 0 {
			parts = append(parts, "fill="+html.EscapeString(rule.Fill.Color[0]))
		}
		if rule.Font != nil && rule.Font.Color != nil {
			parts = append(parts, "font="+html.EscapeString(*rule.Font.Color))
		}
		parts = append(parts, fmt.Sprintf("priority=%d", rule.Priority))
		if rule.StopIfTrue {
			parts = append(parts, "stopIfTrue")
		}
		result += fmt.Sprintf("<li>%s</li>\n", strings.Join(parts, " "))
	}
	result += "</ul>\n"
	result += "<h2>Metadata</h2>\n<ul>\n"
	result += fmt.Sprintf("<li>backend: %s</li>\n", workbook.GetBackendName())
	result += fmt.Sprintf("<li>sheet name: %s</li>\n", html.EscapeString(sheetName))
	result += fmt.Sprintf("<li>range: %s</li>\n", html.EscapeString(rangeStr))
	result += "</ul>\n"
	return mcp.NewToolResultText(result), nil
}
