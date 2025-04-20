package excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

// PagingStrategy はページング範囲の計算戦略を定義するインターフェース
type PagingStrategy interface {
	// CalculatePagingRanges は利用可能なページング範囲のリストを返す
	CalculatePagingRanges() []string
	// ValidatePagingRange は指定された範囲が有効かどうかを検証する
	ValidatePagingRange(rangeStr string) error
}

// ExcelizeFixedSizePagingStrategy は固定サイズでページング範囲を計算する戦略
type ExcelizeFixedSizePagingStrategy struct {
	pageSize  int
	worksheet *ExcelizeWorksheet
	dimension string
}

// NewExcelizeFixedSizePagingStrategy は新しいFixedSizePagingStrategyインスタンスを生成する
func NewExcelizeFixedSizePagingStrategy(pageSize int, worksheet *ExcelizeWorksheet) (*ExcelizeFixedSizePagingStrategy, error) {
	if pageSize <= 0 {
		pageSize = 5000 // デフォルト値
	}

	// シートの次元情報を取得
	dimension, err := worksheet.GetDimention()
	if err != nil {
		return nil, err
	}

	return &ExcelizeFixedSizePagingStrategy{
		pageSize:  pageSize,
		worksheet: worksheet,
		dimension: dimension,
	}, nil
}

// CalculatePagingRanges は固定サイズに基づいてページング範囲のリストを生成する
func (s *ExcelizeFixedSizePagingStrategy) CalculatePagingRanges() []string {
	startCol, startRow, endCol, endRow, err := ParseRange(s.dimension)
	if err != nil {
		return []string{}
	}

	totalCols := endCol - startCol + 1
	cellsPerPage := s.pageSize

	// 1ページあたりの行数を計算
	rowsPerPage := cellsPerPage / totalCols
	if rowsPerPage < 1 {
		rowsPerPage = 1
	}

	var ranges []string
	currentRow := startRow
	for currentRow <= endRow {
		pageEndRow := currentRow + rowsPerPage - 1
		if pageEndRow > endRow {
			pageEndRow = endRow
		}

		startRange, _ := excelize.CoordinatesToCellName(startCol, currentRow)
		endRange, _ := excelize.CoordinatesToCellName(endCol, pageEndRow)
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))

		currentRow = pageEndRow + 1
	}

	return ranges
}

// ValidatePagingRange は指定された範囲が有効かどうかを検証する
func (s *ExcelizeFixedSizePagingStrategy) ValidatePagingRange(rangeStr string) error {
	startCol, startRow, endCol, endRow, err := ParseRange(rangeStr)
	if err != nil {
		return fmt.Errorf("invalid range format: %v", err)
	}

	dimStartCol, dimStartRow, dimEndCol, dimEndRow, err := ParseRange(s.dimension)
	if err != nil {
		return fmt.Errorf("invalid dimension format: %v", err)
	}

	// 範囲がシートの次元内に収まっているか確認
	if startCol < dimStartCol || startRow < dimStartRow ||
		endCol > dimEndCol || endRow > dimEndRow {
		return fmt.Errorf("range %s is outside sheet dimensions %s",
			rangeStr, s.dimension)
	}

	// セル数が pageSize を超えていないか確認
	cellCount := (endRow - startRow + 1) * (endCol - startCol + 1)
	if cellCount > s.pageSize {
		return fmt.Errorf("range contains %d cells, exceeding page size of %d",
			cellCount, s.pageSize)
	}

	return nil
}

func NewOlePagingStrategy(pageSize int, worksheet *OleWorksheet) (PagingStrategy, error) {
	if worksheet == nil {
		return nil, fmt.Errorf("worksheet is nil")
	}

	printAreaPagingStrategy, err := NewPrintAreaPagingStrategy(worksheet)
	if err != nil {
		return nil, err
	}
	printArea, err := printAreaPagingStrategy.getPrintArea()
	if err != nil {
		return nil, err
	}
	if printArea == "" {
		return NewGoxcelFixedSizePagingStrategy(pageSize, worksheet)
	} else {
		return printAreaPagingStrategy, nil
	}
}

// OleFixedSizePagingStrategy は Goxcel を使用した固定サイズでページング範囲を計算する戦略
type OleFixedSizePagingStrategy struct {
	pageSize  int
	worksheet *OleWorksheet
	dimension string
}

// NewGoxcelFixedSizePagingStrategy は新しい GoxcelFixedSizePagingStrategy インスタンスを生成する
func NewGoxcelFixedSizePagingStrategy(pageSize int, worksheet *OleWorksheet) (*OleFixedSizePagingStrategy, error) {
	if pageSize <= 0 {
		pageSize = 5000 // デフォルト値
	}

	if worksheet == nil {
		return nil, fmt.Errorf("worksheet is nil")
	}

	// UsedRange を使用して使用範囲を取得
	dimention, err := worksheet.GetDimention()
	if err != nil {
		return nil, fmt.Errorf("failed to get dimention: %w", err)
	}

	return &OleFixedSizePagingStrategy{
		pageSize:  pageSize,
		worksheet: worksheet,
		dimension: dimention,
	}, nil
}

// CalculatePagingRanges は固定サイズに基づいてページング範囲のリストを生成する
func (s *OleFixedSizePagingStrategy) CalculatePagingRanges() []string {
	startCol, startRow, endCol, endRow, err := ParseRange(s.dimension)
	if err != nil {
		return []string{}
	}

	totalCols := endCol - startCol + 1
	cellsPerPage := s.pageSize

	// 1ページあたりの行数を計算
	rowsPerPage := cellsPerPage / totalCols
	if rowsPerPage < 1 {
		rowsPerPage = 1
	}

	var ranges []string
	currentRow := startRow
	for currentRow <= endRow {
		pageEndRow := currentRow + rowsPerPage - 1
		if pageEndRow > endRow {
			pageEndRow = endRow
		}

		startRange, _ := excelize.CoordinatesToCellName(startCol, currentRow)
		endRange, _ := excelize.CoordinatesToCellName(endCol, pageEndRow)
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))

		currentRow = pageEndRow + 1
	}

	return ranges
}

// ValidatePagingRange は指定された範囲が有効かどうかを検証する
func (s *OleFixedSizePagingStrategy) ValidatePagingRange(rangeStr string) error {
	startCol, startRow, endCol, endRow, err := ParseRange(rangeStr)
	if err != nil {
		return fmt.Errorf("invalid range format: %v", err)
	}

	dimStartCol, dimStartRow, dimEndCol, dimEndRow, err := ParseRange(s.dimension)
	if err != nil {
		return fmt.Errorf("invalid dimension format: %v", err)
	}

	// 範囲がシートの次元内に収まっているか確認
	if startCol < dimStartCol || startRow < dimStartRow ||
		endCol > dimEndCol || endRow > dimEndRow {
		return fmt.Errorf("range %s is outside sheet dimensions %s",
			rangeStr, s.dimension)
	}

	// セル数が pageSize を超えていないか確認
	cellCount := (endRow - startRow + 1) * (endCol - startCol + 1)
	if cellCount > s.pageSize {
		return fmt.Errorf("range contains %d cells, exceeding page size of %d",
			cellCount, s.pageSize)
	}

	return nil
}

// PrintAreaPagingStrategy は印刷範囲とページ区切りに基づいてページング範囲を計算する戦略
type PrintAreaPagingStrategy struct {
	worksheet *OleWorksheet
}

// NewPrintAreaPagingStrategy は新しいPrintAreaPagingStrategyインスタンスを生成する
func NewPrintAreaPagingStrategy(worksheet *OleWorksheet) (*PrintAreaPagingStrategy, error) {
	if worksheet == nil {
		return nil, fmt.Errorf("worksheet is nil")
	}
	return &PrintAreaPagingStrategy{
		worksheet: worksheet,
	}, nil
}

// getPrintArea は印刷範囲を取得する
func (s *PrintAreaPagingStrategy) getPrintArea() (string, error) {
	return s.worksheet.PrintArea()
}

// getHPageBreaksPositions はページ区切りの位置情報を取得する
func (s *PrintAreaPagingStrategy) getHPageBreaksPositions() ([]int, error) {
	pageBreaks, err := s.worksheet.HPageBreaks()
	if err != nil {
		return nil, fmt.Errorf("failed to get HPageBreaks: %w", err)
	}
	return pageBreaks, nil
}

// calculateRangesFromBreaks は印刷範囲とページ区切りから範囲のリストを生成する
func (s *PrintAreaPagingStrategy) calculateRangesFromBreaks(printArea string, breaks []int) []string {
	if printArea == "" {
		return []string{}
	}

	startCol, startRow, endCol, endRow, err := ParseRange(printArea)
	if err != nil {
		return []string{}
	}

	ranges := make([]string, 0)
	currentRow := startRow

	// ページ区切りがない場合は印刷範囲全体を1つの範囲として扱う
	if len(breaks) == 0 {
		startRange, _ := excelize.CoordinatesToCellName(startCol, startRow)
		endRange, _ := excelize.CoordinatesToCellName(endCol, endRow)
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))
		return ranges
	}

	// ページ区切りで範囲を分割
	for _, breakRow := range breaks {
		if breakRow <= startRow || breakRow > endRow {
			continue
		}

		startRange, _ := excelize.CoordinatesToCellName(startCol, currentRow)
		endRange, _ := excelize.CoordinatesToCellName(endCol, breakRow-1)
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))

		currentRow = breakRow
	}

	// 最後の範囲を追加
	if currentRow <= endRow {
		startRange, _ := excelize.CoordinatesToCellName(startCol, currentRow)
		endRange, _ := excelize.CoordinatesToCellName(endCol, endRow)
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))
	}

	return ranges
}

// CalculatePagingRanges は印刷範囲とページ区切りに基づいてページング範囲のリストを生成する
func (s *PrintAreaPagingStrategy) CalculatePagingRanges() []string {
	printArea, err := s.getPrintArea()
	if err != nil {
		return []string{}
	}

	breaks, err := s.getHPageBreaksPositions()
	if err != nil {
		return []string{}
	}

	return s.calculateRangesFromBreaks(printArea, breaks)
}

// ValidatePagingRange は指定された範囲が印刷範囲内に収まっているか検証する
func (s *PrintAreaPagingStrategy) ValidatePagingRange(rangeStr string) error {
	printArea, err := s.getPrintArea()
	if err != nil {
		return fmt.Errorf("failed to get print area: %w", err)
	}
	if printArea == "" {
		return fmt.Errorf("print area is not set")
	}

	// 印刷範囲を解析
	paStartCol, paStartRow, paEndCol, paEndRow, err := ParseRange(printArea)
	if err != nil {
		return fmt.Errorf("invalid print area format: %w", err)
	}

	// 検証対象の範囲を解析
	startCol, startRow, endCol, endRow, err := ParseRange(rangeStr)
	if err != nil {
		return fmt.Errorf("invalid range format: %w", err)
	}

	// 範囲が印刷範囲内に収まっているか確認
	if startCol < paStartCol || startRow < paStartRow ||
		endCol > paEndCol || endRow > paEndRow {
		return fmt.Errorf("range %s is outside print area %s",
			rangeStr, printArea)
	}

	// ページ区切り位置と一致するか確認
	ranges := s.CalculatePagingRanges()
	for _, r := range ranges {
		if r == rangeStr {
			return nil
		}
	}

	return fmt.Errorf("range %s does not match any page break boundary", rangeStr)
}

// PagingRangeService はページング処理を提供するサービス
type PagingRangeService struct {
	strategy PagingStrategy
}

// NewPagingRangeService は新しいPagingRangeServiceインスタンスを生成する
func NewPagingRangeService(strategy PagingStrategy) *PagingRangeService {
	return &PagingRangeService{strategy: strategy}
}

// GetPagingRanges は利用可能なページング範囲のリストを返す
func (s *PagingRangeService) GetPagingRanges() []string {
	return s.strategy.CalculatePagingRanges()
}

// ValidatePagingRange は指定された範囲が有効かどうかを検証する
func (s *PagingRangeService) ValidatePagingRange(rangeStr string) error {
	return s.strategy.ValidatePagingRange(rangeStr)
}

// FilterRemainingPagingRanges は未読の範囲のリストを返す
func (s *PagingRangeService) FilterRemainingPagingRanges(allRanges []string, knownRanges []string) []string {
	if len(knownRanges) == 0 {
		return allRanges
	}

	// knownRanges をマップに変換
	knownMap := make(map[string]bool)
	for _, r := range knownRanges {
		knownMap[r] = true
	}

	// 未読の範囲を抽出
	remaining := make([]string, 0)
	for _, r := range allRanges {
		if !knownMap[r] {
			remaining = append(remaining, r)
		}
	}

	return remaining
}
