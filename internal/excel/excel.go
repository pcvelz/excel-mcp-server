package excel

import (
	"github.com/xuri/excelize/v2"
)

type Excel interface {
	// GetBackendName returns the backend used to manipulate the Excel file.
	GetBackendName() string
	// GetSheets returns a list of all worksheets in the Excel file.
	GetSheets() ([]Worksheet, error)
	// FindSheet finds a sheet by its name and returns a Worksheet.
	FindSheet(sheetName string) (Worksheet, error)
	// CreateNewSheet creates a new sheet with the specified name.
	CreateNewSheet(sheetName string) error
	// CopySheet copies a sheet from one to another.
	CopySheet(srcSheetName, destSheetName string) error
	// SheetNames returns the names of all worksheets in workbook order.
	SheetNames() ([]string, error)
	// RenameSheet renames a sheet, keeping its content and formatting intact.
	// Returns non-fatal warnings about anything the backend could not adjust.
	RenameSheet(oldSheetName, newSheetName string) ([]string, error)
	// DeleteSheet deletes a sheet. It refuses to delete the last remaining
	// sheet, and refuses to leave dangling references behind unless force is
	// set. Returns non-fatal warnings about what was cleaned up.
	DeleteSheet(sheetName string, force bool) ([]string, error)
	// MoveSheet moves a sheet to the given zero-based position in the workbook.
	MoveSheet(sheetName string, index int) error
	// Save saves the Excel file.
	Save() error
}

type Worksheet interface {
	// Release releases the worksheet resources.
	Release()
	// Name returns the name of the worksheet.
	Name() (string, error)
	// GetTable returns a tables in this worksheet.
	GetTables() ([]Table, error)
	// GetPivotTable returns a pivot tables in this worksheet.
	GetPivotTables() ([]PivotTable, error)
	// SetValue sets a value in the specified cell.
	SetValue(cell string, value any) error
	// SetFormula sets a formula in the specified cell.
	SetFormula(cell string, formula string) error
	// GetValue gets the value from the specified cell.
	GetValue(cell string) (string, error)
	// GetRawValue gets the raw (unformatted) value from the specified cell.
	GetRawValue(cell string) (string, error)
	// GetCellType gets the type of the specified cell (number, string, date, bool, formula, error).
	GetCellType(cell string) (string, error)
	// GetFormula gets the formula from the specified cell.
	GetFormula(cell string) (string, error)
	// GetDimention gets the dimension of the worksheet.
	GetDimention() (string, error)
	// GetPagingStrategy returns the paging strategy for the worksheet.
	// The pageSize parameter is used to determine the max size of each page.
	GetPagingStrategy(pageSize int) (PagingStrategy, error)
	// CapturePicture returns base64 encoded image data of the specified range.
	CapturePicture(captureRange string) (string, error)
	// AddTable adds a table to this worksheet.
	AddTable(tableRange, tableName string) error
	// GetCellStyle gets style information for the specified cell.
	GetCellStyle(cell string) (*CellStyle, error)
	// SetCellStyle sets style for the specified cell.
	SetCellStyle(cell string, style *CellStyle) error
	// GetMergedCells returns the merged cell ranges of this worksheet.
	GetMergedCells() ([]string, error)
	// GetColumnWidths returns the explicitly set column widths of this
	// worksheet, keyed by column name (e.g. "B").
	GetColumnWidths(startCol, endCol int) (map[string]float64, error)
	// DeleteRows removes rows startRow..endRow (one-based, inclusive) and
	// shifts everything below them up.
	DeleteRows(startRow, endRow int) error
	// InsertRows inserts count empty rows before beforeRow (one-based).
	InsertRows(beforeRow, count int) error
	// GetConditionalFormatRanges returns the ranges that carry conditional
	// formatting rules on this worksheet.
	GetConditionalFormatRanges() ([]string, error)
	// GetDataValidationRanges returns the ranges that carry data validation
	// rules on this worksheet.
	GetDataValidationRanges() ([]string, error)
	// GetConditionalFormats returns every conditional formatting rule on this
	// worksheet, ordered by the priority Excel evaluates them in.
	//
	// A flat slice rather than a map keyed by range: a worksheet may hold
	// several <conditionalFormatting> blocks sharing one sqref, which Excel
	// writes routinely, so keying by range silently drops all but the last.
	GetConditionalFormats() ([]ConditionalFormatRule, error)
	// SetConditionalFormat replaces the conditional formatting on rangeRef
	// with the given rules, which are applied in the order supplied.
	SetConditionalFormat(rangeRef string, rules []ConditionalFormatRule) error
	// ClearConditionalFormat removes the conditional formatting on rangeRef.
	ClearConditionalFormat(rangeRef string) error
}

// ConditionalFormatRule is one conditional formatting rule, as Excel stores
// it. The fields mirror a <cfRule> element rather than any excelize type,
// because the parts that decide what a sheet actually looks like -- priority,
// stopIfTrue, and the differential style the rule applies -- are not all
// reachable through excelize's exported API.
type ConditionalFormatRule struct {
	// Range is the sqref the owning block applies to, e.g. "A1:A11" or the
	// multi-area "B2:B9 D2:D9".
	Range string `yaml:"range"`
	// Type is the raw OOXML rule type: cellIs, expression, colorScale,
	// dataBar, iconSet, timePeriod, containsText, top10, aboveAverage, ...
	Type string `yaml:"type"`
	// Operator qualifies a cellIs or text rule: lessThanOrEqual, equal,
	// between, greaterThan, ... Empty for rules that take no operator.
	Operator string `yaml:"operator,omitempty"`
	// Formulas holds the <formula> children in document order. A cellIs rule
	// has one, or two for between/notBetween; an expression rule has the
	// single custom formula.
	Formulas []string `yaml:"formulas,omitempty"`
	// Text is the needle for containsText and its siblings.
	Text string `yaml:"text,omitempty"`
	// Priority decides which rule wins where ranges overlap: lower first.
	Priority int `yaml:"priority"`
	// StopIfTrue halts evaluation of lower-priority rules once this matches.
	StopIfTrue bool `yaml:"stopIfTrue,omitempty"`
	// Font and Fill are the differential style (dxf) applied on a match. Fill
	// is the background colour, which is what most rules exist for.
	Font *FontStyle `yaml:"font,omitempty"`
	Fill *FillStyle `yaml:"fill,omitempty"`
	// Colors carries the stop colours of a colorScale, or the bar colour of a
	// dataBar -- gradient colouring keeps its colours here, not in Fill.
	Colors []string `yaml:"colors,omitempty"`
	// Thresholds are the colorScale/dataBar/iconSet value objects (cfvo), in
	// order, rendered as "type" or "type=value" (e.g. "min", "percentile=50").
	Thresholds []string `yaml:"thresholds,omitempty"`
	// StyleError explains why the applied style could not be resolved, for a
	// rule that references a differential style the workbook never defines.
	// Excel renders such a rule as no formatting at all.
	StyleError string `yaml:"styleError,omitempty"`
}

type Table struct {
	Name  string
	Range string
}

type PivotTable struct {
	Name  string
	Range string
}

type CellStyle struct {
	Border        []Border        `yaml:"border,omitempty"`
	Font          *FontStyle      `yaml:"font,omitempty"`
	Fill          *FillStyle      `yaml:"fill,omitempty"`
	Alignment     *AlignmentStyle `yaml:"alignment,omitempty"`
	NumFmt        *string         `yaml:"numFmt,omitempty"`
	DecimalPlaces *int            `yaml:"decimalPlaces,omitempty"`
}

type AlignmentStyle struct {
	Horizontal   *string `yaml:"horizontal,omitempty"` // left, center, right, fill, justify, centerContinuous, distributed
	Vertical     *string `yaml:"vertical,omitempty"`   // top, center, bottom, justify, distributed
	WrapText     *bool   `yaml:"wrapText,omitempty"`
	ShrinkToFit  *bool   `yaml:"shrinkToFit,omitempty"`
	TextRotation *int    `yaml:"textRotation,omitempty"` // 0-180 or 255 for vertical
	Indent       *int    `yaml:"indent,omitempty"`
}

type Border struct {
	Type  BorderType  `yaml:"type"`
	Style BorderStyle `yaml:"style,omitempty"`
	Color string      `yaml:"color,omitempty"`
}

type FontStyle struct {
	Bold      *bool          `yaml:"bold,omitempty"`
	Italic    *bool          `yaml:"italic,omitempty"`
	Underline *FontUnderline `yaml:"underline,omitempty"`
	Size      *int           `yaml:"size,omitempty"`
	Strike    *bool          `yaml:"strike,omitempty"`
	Color     *string        `yaml:"color,omitempty"`
	VertAlign *FontVertAlign `yaml:"vertAlign,omitempty"`
}

type FillStyle struct {
	Type    FillType     `yaml:"type,omitempty"`
	Pattern FillPattern  `yaml:"pattern,omitempty"`
	Color   []string     `yaml:"color,omitempty"`
	Shading *FillShading `yaml:"shading,omitempty"`
}

// OpenFile opens an Excel file and returns an Excel interface.
// It first tries to open the file using OLE automation, and if that fails,
// it tries to using the excelize library.
func OpenFile(absoluteFilePath string) (Excel, func(), error) {
	ole, releaseFn, err := NewExcelOle(absoluteFilePath)
	if err == nil {
		return ole, releaseFn, nil
	}
	// If OLE fails, try Excelize
	workbook, err := excelize.OpenFile(absoluteFilePath)
	if err != nil {
		return nil, func() {}, err
	}
	excelize := NewExcelizeExcel(workbook)
	return excelize, func() {
		workbook.Close()
	}, nil
}

// BorderType represents border direction
type BorderType string

const (
	BorderTypeLeft         BorderType = "left"
	BorderTypeRight        BorderType = "right"
	BorderTypeTop          BorderType = "top"
	BorderTypeBottom       BorderType = "bottom"
	BorderTypeDiagonalDown BorderType = "diagonalDown"
	BorderTypeDiagonalUp   BorderType = "diagonalUp"
)

func (b BorderType) String() string {
	return string(b)
}

func (b BorderType) MarshalText() ([]byte, error) {
	return []byte(b.String()), nil
}

func BorderTypeValues() []BorderType {
	return []BorderType{
		BorderTypeLeft,
		BorderTypeRight,
		BorderTypeTop,
		BorderTypeBottom,
		BorderTypeDiagonalDown,
		BorderTypeDiagonalUp,
	}
}

// BorderStyle represents border style constants
type BorderStyle string

const (
	BorderStyleNone             BorderStyle = "none"
	BorderStyleContinuous       BorderStyle = "continuous"
	BorderStyleDash             BorderStyle = "dash"
	BorderStyleDot              BorderStyle = "dot"
	BorderStyleDouble           BorderStyle = "double"
	BorderStyleDashDot          BorderStyle = "dashDot"
	BorderStyleDashDotDot       BorderStyle = "dashDotDot"
	BorderStyleSlantDashDot     BorderStyle = "slantDashDot"
	BorderStyleMediumDashDot    BorderStyle = "mediumDashDot"
	BorderStyleMediumDashDotDot BorderStyle = "mediumDashDotDot"
)

func (b BorderStyle) String() string {
	return string(b)
}

func (b BorderStyle) MarshalText() ([]byte, error) {
	return []byte(b.String()), nil
}

func BorderStyleValues() []BorderStyle {
	return []BorderStyle{
		BorderStyleNone,
		BorderStyleContinuous,
		BorderStyleDash,
		BorderStyleDot,
		BorderStyleDouble,
		BorderStyleDashDot,
		BorderStyleDashDotDot,
		BorderStyleSlantDashDot,
		BorderStyleMediumDashDot,
		BorderStyleMediumDashDotDot,
	}
}

// FontUnderline represents underline styles for font
type FontUnderline string

const (
	FontUnderlineNone             FontUnderline = "none"
	FontUnderlineSingle           FontUnderline = "single"
	FontUnderlineDouble           FontUnderline = "double"
	FontUnderlineSingleAccounting FontUnderline = "singleAccounting"
	FontUnderlineDoubleAccounting FontUnderline = "doubleAccounting"
)

func (f FontUnderline) String() string {
	return string(f)
}
func (f FontUnderline) MarshalText() ([]byte, error) {
	return []byte(f.String()), nil
}

func FontUnderlineValues() []FontUnderline {
	return []FontUnderline{
		FontUnderlineNone,
		FontUnderlineSingle,
		FontUnderlineDouble,
		FontUnderlineSingleAccounting,
		FontUnderlineDoubleAccounting,
	}
}

// FontVertAlign represents vertical alignment options for font styles
type FontVertAlign string

const (
	FontVertAlignBaseline    FontVertAlign = "baseline"
	FontVertAlignSuperscript FontVertAlign = "superscript"
	FontVertAlignSubscript   FontVertAlign = "subscript"
)

func (v FontVertAlign) String() string {
	return string(v)
}

func (v FontVertAlign) MarshalText() ([]byte, error) {
	return []byte(v.String()), nil
}

func FontVertAlignValues() []FontVertAlign {
	return []FontVertAlign{
		FontVertAlignBaseline,
		FontVertAlignSuperscript,
		FontVertAlignSubscript,
	}
}

// FillType represents fill types for cell styles
type FillType string

const (
	FillTypeGradient FillType = "gradient"
	FillTypePattern  FillType = "pattern"
)

func (f FillType) String() string {
	return string(f)
}

func (f FillType) MarshalText() ([]byte, error) {
	return []byte(f.String()), nil
}

func FillTypeValues() []FillType {
	return []FillType{
		FillTypeGradient,
		FillTypePattern,
	}
}

// FillPattern represents fill pattern constants
type FillPattern string

const (
	FillPatternNone            FillPattern = "none"
	FillPatternSolid           FillPattern = "solid"
	FillPatternMediumGray      FillPattern = "mediumGray"
	FillPatternDarkGray        FillPattern = "darkGray"
	FillPatternLightGray       FillPattern = "lightGray"
	FillPatternDarkHorizontal  FillPattern = "darkHorizontal"
	FillPatternDarkVertical    FillPattern = "darkVertical"
	FillPatternDarkDown        FillPattern = "darkDown"
	FillPatternDarkUp          FillPattern = "darkUp"
	FillPatternDarkGrid        FillPattern = "darkGrid"
	FillPatternDarkTrellis     FillPattern = "darkTrellis"
	FillPatternLightHorizontal FillPattern = "lightHorizontal"
	FillPatternLightVertical   FillPattern = "lightVertical"
	FillPatternLightDown       FillPattern = "lightDown"
	FillPatternLightUp         FillPattern = "lightUp"
	FillPatternLightGrid       FillPattern = "lightGrid"
	FillPatternLightTrellis    FillPattern = "lightTrellis"
	FillPatternGray125         FillPattern = "gray125"
	FillPatternGray0625        FillPattern = "gray0625"
)

func (f FillPattern) String() string {
	return string(f)
}

func (f FillPattern) MarshalText() ([]byte, error) {
	return []byte(f.String()), nil
}

func FillPatternValues() []FillPattern {
	return []FillPattern{
		FillPatternNone,
		FillPatternSolid,
		FillPatternMediumGray,
		FillPatternDarkGray,
		FillPatternLightGray,
		FillPatternDarkHorizontal,
		FillPatternDarkVertical,
		FillPatternDarkDown,
		FillPatternDarkUp,
		FillPatternDarkGrid,
		FillPatternDarkTrellis,
		FillPatternLightHorizontal,
		FillPatternLightVertical,
		FillPatternLightDown,
		FillPatternLightUp,
		FillPatternLightGrid,
		FillPatternLightTrellis,
		FillPatternGray125,
		FillPatternGray0625,
	}
}

// FillShading represents fill shading constants
type FillShading string

const (
	FillShadingHorizontal   FillShading = "horizontal"
	FillShadingVertical     FillShading = "vertical"
	FillShadingDiagonalDown FillShading = "diagonalDown"
	FillShadingDiagonalUp   FillShading = "diagonalUp"
	FillShadingFromCenter   FillShading = "fromCenter"
	FillShadingFromCorner   FillShading = "fromCorner"
)

func (f FillShading) String() string {
	return string(f)
}

func (f FillShading) MarshalText() ([]byte, error) {
	return []byte(f.String()), nil
}

func FillShadingValues() []FillShading {
	return []FillShading{
		FillShadingHorizontal,
		FillShadingVertical,
		FillShadingDiagonalDown,
		FillShadingDiagonalUp,
		FillShadingFromCenter,
		FillShadingFromCorner,
	}
}
