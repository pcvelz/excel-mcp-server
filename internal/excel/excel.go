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
	Border        []BorderStyle `yaml:"border,omitempty"`
	Font          *FontStyle    `yaml:"font,omitempty"`
	Fill          *FillStyle    `yaml:"fill,omitempty"`
	NumFmt        string        `yaml:"numFmt,omitempty"`
	DecimalPlaces int           `yaml:"decimalPlaces,omitempty"`
}

type BorderStyle struct {
	Type  string          `yaml:"type"`
	Style BorderStyleName `yaml:"style,omitempty"`
	Color string          `yaml:"color,omitempty"`
}

type FontStyle struct {
	Bold      bool   `yaml:"bold,omitempty"`
	Italic    bool   `yaml:"italic,omitempty"`
	Underline string `yaml:"underline,omitempty"`
	Size      int    `yaml:"size,omitempty"`
	Strike    bool   `yaml:"strike,omitempty"`
	Color     string `yaml:"color,omitempty"`
	VertAlign string `yaml:"vertAlign,omitempty"`
}

type FillStyle struct {
	Type    string          `yaml:"type,omitempty"`
	Pattern FillPatternName `yaml:"pattern,omitempty"`
	Color   []string        `yaml:"color,omitempty"`
	Shading FillShadingName `yaml:"shading,omitempty"`
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

// BorderStyleName represents border style constants
type BorderStyleName int

const (
	BorderStyleNone BorderStyleName = iota
	BorderStyleContinuous
	BorderStyleDash
	BorderStyleDot
	BorderStyleDouble
	BorderStyleDashDot
	BorderStyleDashDotDot
	BorderStyleSlantDashDot
	BorderStyleMediumDashDot
	BorderStyleMediumDashDotDot
)

var borderStyleNames = map[BorderStyleName]string{
	BorderStyleNone:             "none",
	BorderStyleContinuous:       "continuous",
	BorderStyleDash:             "dash",
	BorderStyleDot:              "dot",
	BorderStyleDouble:           "double",
	BorderStyleDashDot:          "dashDot",
	BorderStyleDashDotDot:       "dashDotDot",
	BorderStyleSlantDashDot:     "slantDashDot",
	BorderStyleMediumDashDot:    "mediumDashDot",
	BorderStyleMediumDashDotDot: "mediumDashDotDot",
}

func (b BorderStyleName) String() string {
	if name, exists := borderStyleNames[b]; exists {
		return name
	}
	return "continuous"
}

func (b BorderStyleName) MarshalText() ([]byte, error) {
	return []byte(b.String()), nil
}

// FillPatternName represents fill pattern constants
type FillPatternName int

const (
	FillPatternNone FillPatternName = iota
	FillPatternSolid
	FillPatternMediumGray
	FillPatternDarkGray
	FillPatternLightGray
	FillPatternDarkHorizontal
	FillPatternDarkVertical
	FillPatternDarkDown
	FillPatternDarkUp
	FillPatternDarkGrid
	FillPatternDarkTrellis
	FillPatternLightHorizontal
	FillPatternLightVertical
	FillPatternLightDown
	FillPatternLightUp
	FillPatternLightGrid
	FillPatternLightTrellis
	FillPatternGray125
	FillPatternGray0625
)

var fillPatternNames = map[FillPatternName]string{
	FillPatternNone:            "none",
	FillPatternSolid:           "solid",
	FillPatternMediumGray:      "mediumGray",
	FillPatternDarkGray:        "darkGray",
	FillPatternLightGray:       "lightGray",
	FillPatternDarkHorizontal:  "darkHorizontal",
	FillPatternDarkVertical:    "darkVertical",
	FillPatternDarkDown:        "darkDown",
	FillPatternDarkUp:          "darkUp",
	FillPatternDarkGrid:        "darkGrid",
	FillPatternDarkTrellis:     "darkTrellis",
	FillPatternLightHorizontal: "lightHorizontal",
	FillPatternLightVertical:   "lightVertical",
	FillPatternLightDown:       "lightDown",
	FillPatternLightUp:         "lightUp",
	FillPatternLightGrid:       "lightGrid",
	FillPatternLightTrellis:    "lightTrellis",
	FillPatternGray125:         "gray125",
	FillPatternGray0625:        "gray0625",
}

func (f FillPatternName) String() string {
	if name, exists := fillPatternNames[f]; exists {
		return name
	}
	return "none"
}

func (f FillPatternName) MarshalText() ([]byte, error) {
	return []byte(f.String()), nil
}

// FillShadingName represents fill shading constants
type FillShadingName int

const (
	FillShadingHorizontal FillShadingName = iota
	FillShadingVertical
	FillShadingDiagonalDown
	FillShadingDiagonalUp
	FillShadingFromCenter
	FillShadingFromCorner
)

func (f FillShadingName) String() string {
	if name, exists := fillShadingNames[f]; exists {
		return name
	}
	return "horizontal"
}

func (f FillShadingName) MarshalText() ([]byte, error) {
	return []byte(f.String()), nil
}

var fillShadingNames = map[FillShadingName]string{
	FillShadingHorizontal:   "horizontal",
	FillShadingVertical:     "vertical",
	FillShadingDiagonalDown: "diagonalDown",
	FillShadingDiagonalUp:   "diagonalUp",
	FillShadingFromCenter:   "fromCenter",
	FillShadingFromCorner:   "fromCorner",
}
