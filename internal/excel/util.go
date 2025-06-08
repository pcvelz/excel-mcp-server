package excel

import (
	"fmt"
	"os"
	"path"
	"regexp"
	"time"

	"github.com/xuri/excelize/v2"
)

// parseRange parses Excel's range string (e.g. A1:C10)
func ParseRange(rangeStr string) (int, int, int, int, error) {
	re := regexp.MustCompile(`(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)`)
	matches := re.FindStringSubmatch(rangeStr)
	if matches == nil {
		return 0, 0, 0, 0, fmt.Errorf("invalid range format: %s", rangeStr)
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(matches[1])
	if err != nil {
		return 0, 0, 0, 0, err
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

// FetchLastModifiedTimeWithExcelize retrieves the last modified time of an Excel file using excelize
func FetchLastModifiedTimeWithExcelize(absolutePath string) (time.Time, error) {
	f, err := excelize.OpenFile(absolutePath)
	if err != nil {
		return time.Time{}, err
	}
	defer f.Close()
	props, err := f.GetDocProps()
	if err != nil {
		return time.Time{}, err
	}
	lastSaveTime, err := time.Parse(time.RFC3339, props.Modified)
	if err != nil {
		return time.Time{}, err
	}
	return lastSaveTime.Local(), nil
}

func FileIsNotWritable(absolutePath string) bool {
	f, err := os.OpenFile(path.Clean(absolutePath), os.O_WRONLY, os.ModePerm)
	if err != nil {
		return true
	}
	defer f.Close()
	return false
}
