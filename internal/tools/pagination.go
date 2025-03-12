package tools

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

// FixedSizePagingStrategy は固定サイズでページング範囲を計算する戦略
type FixedSizePagingStrategy struct {
	pageSize  int
	workbook  *excelize.File
	sheetName string
	dimension string
}

// NewFixedSizePagingStrategy は新しいFixedSizePagingStrategyインスタンスを生成する
func NewFixedSizePagingStrategy(pageSize int, workbook *excelize.File, sheetName string) (*FixedSizePagingStrategy, error) {
	if pageSize <= 0 {
		pageSize = 5000 // デフォルト値
	}

	// シート名の検証
	index, err := workbook.GetSheetIndex(sheetName)
	if err != nil || index == -1 {
		return nil, fmt.Errorf("sheet %s not found", sheetName)
	}

	// シートの次元情報を取得
	dimension, err := workbook.GetSheetDimension(sheetName)
	if err != nil {
		return nil, err
	}

	return &FixedSizePagingStrategy{
		pageSize:  pageSize,
		workbook:  workbook,
		sheetName: sheetName,
		dimension: dimension,
	}, nil
}

// CalculatePagingRanges は固定サイズに基づいてページング範囲のリストを生成する
func (s *FixedSizePagingStrategy) CalculatePagingRanges() []string {
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
		// 範囲文字列を生成
		ranges = append(ranges, fmt.Sprintf("%s:%s", startRange, endRange))

		currentRow = pageEndRow + 1
	}

	return ranges
}

// ValidatePagingRange は指定された範囲が有効かどうかを検証する
func (s *FixedSizePagingStrategy) ValidatePagingRange(rangeStr string) error {
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
