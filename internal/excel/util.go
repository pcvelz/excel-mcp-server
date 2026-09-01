package excel

import (
	"fmt"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"slices"
	"strings"

	"github.com/xuri/excelize/v2"
)

// parseRange parses Excel's range string (e.g. A1:C10 or A1)
func ParseRange(rangeStr string) (int, int, int, int, error) {
	re := regexp.MustCompile(`^(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?$`)
	matches := re.FindStringSubmatch(rangeStr)
	if matches == nil {
		return 0, 0, 0, 0, fmt.Errorf("invalid range format: %s", rangeStr)
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(matches[1])
	if err != nil {
		return 0, 0, 0, 0, err
	}

	if matches[2] == "" {
		// Single cell case
		return startCol, startRow, startCol, startRow, nil
	}

	endCol, endRow, err := excelize.CellNameToCoordinates(matches[2])
	if err != nil {
		return 0, 0, 0, 0, err
	}
	return startCol, startRow, endCol, endRow, nil
}

func NormalizeRange(rangeStr string) string {
	startCol, startRow, endCol, endRow, _ := ParseRange(rangeStr)
	startCell, _ := excelize.CoordinatesToCellName(startCol, startRow)
	endCell, _ := excelize.CoordinatesToCellName(endCol, endRow)
	return fmt.Sprintf("%s:%s", startCell, endCell)
}

// supportedWorkbookExtensions are the formats the excelize backend can write.
var supportedWorkbookExtensions = []string{".xlsx", ".xlsm", ".xltx", ".xltm"}

// CreateWorkbook creates a new, empty workbook at absolutePath holding a single
// sheet with the given name. It refuses to overwrite an existing file.
func CreateWorkbook(absolutePath string, sheetName string) error {
	if _, err := os.Stat(absolutePath); err == nil {
		return fmt.Errorf("file already exists: %s", absolutePath)
	}
	dir := filepath.Dir(absolutePath)
	if info, err := os.Stat(dir); err != nil || !info.IsDir() {
		return fmt.Errorf("cannot create %s: the directory %s does not exist", absolutePath, dir)
	}
	extension := strings.ToLower(filepath.Ext(absolutePath))
	if !slices.Contains(supportedWorkbookExtensions, extension) {
		return fmt.Errorf("cannot create %s: unsupported file extension %q (expected one of %s)",
			absolutePath, extension, strings.Join(supportedWorkbookExtensions, ", "))
	}

	workbook := excelize.NewFile()
	defer workbook.Close()

	defaultSheet := workbook.GetSheetName(0)
	if sheetName != defaultSheet {
		if err := workbook.SetSheetName(defaultSheet, sheetName); err != nil {
			return err
		}
	}
	// Write through os.OpenFile rather than SaveAs: excelize's SaveAs rejects
	// paths longer than 207 characters, a limit this server deliberately
	// ignores (see ExcelizeExcel.Save).
	file, err := os.OpenFile(filepath.Clean(absolutePath), os.O_WRONLY|os.O_TRUNC|os.O_CREATE, 0o644)
	if err != nil {
		return err
	}
	defer file.Close()
	return workbook.Write(file)
}

// FileIsNotReadable checks if a file is not writable
func FileIsNotWritable(absolutePath string) bool {
	f, err := os.OpenFile(path.Clean(absolutePath), os.O_WRONLY, os.ModePerm)
	if err != nil {
		return true
	}
	defer f.Close()
	return false
}
