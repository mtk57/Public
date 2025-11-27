using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SimpleExcelGrep.Services
{
    /// <summary>
    /// 「その他」画面のバッチ処理を担当するサービス
    /// </summary>
    public class ExcelModificationService
    {
        private readonly LogService _logger;

        public ExcelModificationService(LogService logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// 指定されたオプションに従ってExcelファイルを一括処理する
        /// </summary>
        public async Task<OtherOperationResult> RunAsync(OtherOperationOptions options, Action<int, int, string> progressCallback, CancellationToken cancellationToken)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (!options.RemoveAllFormulas && !options.RemoveAllShapes && !options.DeleteByKeywordEnabled) return new OtherOperationResult();

            var result = new OtherOperationResult();
            var searchOption = options.IncludeSubDirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var excelFiles = Directory.GetFiles(options.FolderPath, "*.xlsx", searchOption)
                                      .Concat(Directory.GetFiles(options.FolderPath, "*.xlsm", searchOption))
                                      .ToArray();

            result.TotalFiles = excelFiles.Length;

            var logEnabled = options.EnableLogOutput;
            var deleteByKeywordInfo = options.DeleteByKeywordEnabled
                ? $"削除方向={(options.DeleteByKeywordDirectionIsRow ? "行->列削除" : "列->行削除")}, 全対象={options.DeleteByKeywordTargetAll}, 指定={(options.DeleteByKeywordTargetAll ? "ALL" : string.Join(",", options.DeleteByKeywordTargets.OrderBy(x => x)))}, 完全一致={options.DeleteByKeywordFullMatch}, 大小区別={options.DeleteByKeywordCaseSensitive}, 全半角区別={options.DeleteByKeywordWidthSensitive}, キーワード=\"{string.Join(",", options.DeleteByKeywordKeywords)}\""
                : "削除無効";

            _logger.LogMessage($"その他処理開始: フォルダ={options.FolderPath}, サブフォルダ={options.IncludeSubDirectories}, 無視サイズ(MB)={options.IgnoreFileSizeMb}, 非表示シート対象={options.IncludeInvisibleSheets}, 並列数={options.MaxParallelism}, 図削除={options.RemoveAllShapes}, 数式削除={options.RemoveAllFormulas}, キーワード削除={deleteByKeywordInfo}", force: logEnabled);

            progressCallback?.Invoke(0, result.TotalFiles, "処理開始");

            int processed = 0;
            int modified = 0;
            int errors = 0;

            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = Math.Max(1, options.MaxParallelism),
                CancellationToken = cancellationToken
            };

            await Task.Run(() =>
            {
                Parallel.ForEach(excelFiles, parallelOptions, filePath =>
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (options.IgnoreFileSizeMb > 0)
                    {
                        try
                        {
                            var fileInfo = new FileInfo(filePath);
                            var fileSizeMb = (double)fileInfo.Length / (1024 * 1024);
                            if (fileSizeMb > options.IgnoreFileSizeMb)
                            {
                                _logger.LogMessage($"ファイルサイズ超過のためスキップ: {filePath}", force: logEnabled);
                                var current = Interlocked.Increment(ref processed);
                                progressCallback?.Invoke(current, result.TotalFiles, $"{current}/{result.TotalFiles} 処理済み (スキップ)");
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogMessage($"ファイルサイズ確認でエラー: {filePath}, {ex.GetType().Name}: {ex.Message}", force: true);
                        }
                    }

                    bool fileModified = false;
                    try
                    {
                        using (var document = SpreadsheetDocument.Open(filePath, true))
                        {
                            var workbookPart = document.WorkbookPart;
                            if (workbookPart == null || workbookPart.Workbook == null)
                            {
                                _logger.LogMessage($"ワークブック情報が取得できないためスキップ: {filePath}", force: true);
                            }
                            else
                            {
                                var sharedStrings = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable;

                                if (options.RemoveAllShapes)
                                {
                                    fileModified |= RemoveAllShapes(workbookPart, options.IncludeInvisibleSheets, cancellationToken);
                                }
                                if (options.RemoveAllFormulas)
                                {
                                    fileModified |= ReplaceFormulasWithValues(workbookPart, sharedStrings, options.IncludeInvisibleSheets, cancellationToken);
                                }
                                if (options.DeleteByKeywordEnabled)
                                {
                                    fileModified |= RemoveRowsOrColumnsByKeyword(workbookPart, sharedStrings, options, cancellationToken);
                                }

                                if (fileModified && workbookPart.CalculationChainPart != null)
                                {
                                    workbookPart.DeletePart(workbookPart.CalculationChainPart);
                                }

                                if (fileModified)
                                {
                                    workbookPart.Workbook.Save();
                                }
                            }
                        }

                        if (fileModified)
                        {
                            Interlocked.Increment(ref modified);
                            _logger.LogMessage($"処理完了: {filePath}", force: logEnabled);
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        Interlocked.Increment(ref errors);
                        _logger.LogMessage($"その他処理でエラー: {filePath}, {ex.GetType().Name}: {ex.Message}", force: true);
                    }
                    finally
                    {
                        var current = Interlocked.Increment(ref processed);
                        progressCallback?.Invoke(current, result.TotalFiles, $"{current}/{result.TotalFiles} ファイル処理済み");
                    }
                });
            }, cancellationToken);

            result.ProcessedFiles = processed;
            result.ModifiedFiles = modified;
            result.ErrorCount = errors;

            _logger.LogMessage($"その他処理完了: 対象={result.TotalFiles}件, 処理済み={result.ProcessedFiles}件, 変更あり={result.ModifiedFiles}件, エラー={result.ErrorCount}件", force: logEnabled);
            return result;
        }

        private bool RemoveAllShapes(WorkbookPart workbookPart, bool includeInvisibleSheets, CancellationToken cancellationToken)
        {
            bool changed = false;

            foreach (var worksheetPart in workbookPart.WorksheetParts)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var sheet = GetSheet(workbookPart, worksheetPart);
                if (!includeInvisibleSheets && sheet != null && sheet.State != null && sheet.State.Value != SheetStateValues.Visible)
                {
                    continue;
                }

                bool sheetChanged = false;

                var drawingRef = worksheetPart.Worksheet.Elements<Drawing>().FirstOrDefault();
                if (drawingRef != null)
                {
                    drawingRef.Remove();
                    sheetChanged = true;
                }

                if (worksheetPart.DrawingsPart != null)
                {
                    worksheetPart.DeletePart(worksheetPart.DrawingsPart);
                    sheetChanged = true;
                }

                var legacyDrawing = worksheetPart.Worksheet.Elements<LegacyDrawing>().FirstOrDefault();
                if (legacyDrawing != null)
                {
                    legacyDrawing.Remove();
                    sheetChanged = true;
                }

                var legacyDrawingHf = worksheetPart.Worksheet.Elements<LegacyDrawingHeaderFooter>().FirstOrDefault();
                if (legacyDrawingHf != null)
                {
                    legacyDrawingHf.Remove();
                    sheetChanged = true;
                }

                var vmlParts = worksheetPart.VmlDrawingParts.ToList();
                foreach (var vmlPart in vmlParts)
                {
                    worksheetPart.DeletePart(vmlPart);
                    sheetChanged = true;
                }

                if (sheetChanged)
                {
                    worksheetPart.Worksheet.Save();
                    changed = true;
                }
            }

            return changed;
        }

        private bool ReplaceFormulasWithValues(WorkbookPart workbookPart, SharedStringTable sharedStrings, bool includeInvisibleSheets, CancellationToken cancellationToken)
        {
            bool changed = false;

            foreach (var worksheetPart in workbookPart.WorksheetParts)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var sheet = GetSheet(workbookPart, worksheetPart);
                if (!includeInvisibleSheets && sheet != null && sheet.State != null && sheet.State.Value != SheetStateValues.Visible)
                {
                    continue;
                }

                bool sheetChanged = false;

                foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (cell.CellFormula == null) continue;

                    var displayValue = GetCellDisplayValue(cell, sharedStrings);
                    ApplyValueToCell(cell, displayValue);
                    sheetChanged = true;
                }

                if (sheetChanged)
                {
                    worksheetPart.Worksheet.Save();
                    changed = true;
                }
            }

            return changed;
        }

        private bool RemoveRowsOrColumnsByKeyword(WorkbookPart workbookPart, SharedStringTable sharedStrings, OtherOperationOptions options, CancellationToken cancellationToken)
        {
            bool changed = false;
            var compareOptions = BuildCompareOptions(options.DeleteByKeywordCaseSensitive, options.DeleteByKeywordWidthSensitive);
            var keywords = options.DeleteByKeywordKeywords ?? Array.Empty<string>();

            foreach (var worksheetPart in workbookPart.WorksheetParts)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var sheet = GetSheet(workbookPart, worksheetPart);
                if (!options.IncludeInvisibleSheets && sheet != null && sheet.State != null && sheet.State.Value != SheetStateValues.Visible)
                {
                    continue;
                }

                var worksheet = worksheetPart.Worksheet;
                var sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData == null) continue;

                if (options.DeleteByKeywordDirectionIsRow)
                {
                    var columnsToDelete = new HashSet<int>();

                    foreach (var row in sheetData.Elements<Row>())
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var rowIndex = row.RowIndex != null ? (int)row.RowIndex.Value : 0;
                        if (rowIndex <= 0) continue;

                        if (!options.DeleteByKeywordTargetAll && !options.DeleteByKeywordTargets.Contains(rowIndex))
                        {
                            continue;
                        }

                        foreach (var cell in row.Elements<Cell>())
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (!IsKeywordHit(cell, sharedStrings, keywords, compareOptions, options.DeleteByKeywordFullMatch))
                            {
                                continue;
                            }

                            var columnIndex = GetColumnIndex(cell.CellReference);
                            if (columnIndex.HasValue)
                            {
                                columnsToDelete.Add(columnIndex.Value);
                            }
                        }
                    }

                    if (columnsToDelete.Count > 0)
                    {
                        DeleteColumns(worksheet, columnsToDelete, cancellationToken);
                        changed = true;
                    }
                }
                else
                {
                    var rowsToDelete = new HashSet<int>();

                    foreach (var row in sheetData.Elements<Row>())
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var rowIndex = row.RowIndex != null ? (int)row.RowIndex.Value : 0;
                        if (rowIndex <= 0) continue;

                        foreach (var cell in row.Elements<Cell>())
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            var columnIndex = GetColumnIndex(cell.CellReference);
                            if (!columnIndex.HasValue)
                            {
                                continue;
                            }

                            if (!options.DeleteByKeywordTargetAll && !options.DeleteByKeywordTargets.Contains(columnIndex.Value))
                            {
                                continue;
                            }

                            if (!IsKeywordHit(cell, sharedStrings, keywords, compareOptions, options.DeleteByKeywordFullMatch))
                            {
                                continue;
                            }

                            rowsToDelete.Add(rowIndex);
                        }
                    }

                    if (rowsToDelete.Count > 0)
                    {
                        DeleteRows(worksheet, rowsToDelete, cancellationToken);
                        changed = true;
                    }
                }
            }

            return changed;
        }

        private CompareOptions BuildCompareOptions(bool caseSensitive, bool widthSensitive)
        {
            var options = CompareOptions.None;
            if (!caseSensitive)
            {
                options |= CompareOptions.IgnoreCase;
            }

            if (!widthSensitive)
            {
                options |= CompareOptions.IgnoreWidth;
            }

            return options;
        }

        private bool IsKeywordHit(Cell cell, SharedStringTable sharedStrings, IEnumerable<string> keywords, CompareOptions compareOptions, bool fullMatch)
        {
            var text = GetCellDisplayValue(cell, sharedStrings);
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            var compareInfo = CultureInfo.CurrentCulture.CompareInfo;

            foreach (var keyword in keywords)
            {
                if (string.IsNullOrEmpty(keyword)) continue;

                if (fullMatch)
                {
                    if (compareInfo.Compare(text, keyword, compareOptions) == 0)
                    {
                        return true;
                    }
                }
                else
                {
                    if (compareInfo.IndexOf(text, keyword, compareOptions) >= 0)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private string GetCellDisplayValue(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell == null) return string.Empty;

            string cellValueStr = cell.CellValue?.InnerText ?? string.Empty;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && sharedStringTable != null)
            {
                if (int.TryParse(cellValueStr, out int ssid) && ssid >= 0 && ssid < sharedStringTable.ChildElements.Count)
                {
                    var ssi = sharedStringTable.ChildElements[ssid] as SharedStringItem;
                    if (ssi != null)
                    {
                        return string.Concat(ssi.Elements<Text>().Select(t => t.Text));
                    }
                }
                return string.Empty;
            }

            return cellValueStr;
        }

        private void ApplyValueToCell(Cell cell, string displayValue)
        {
            if (cell == null) return;

            cell.CellFormula = null;

            if (cell.CellValue == null)
            {
                cell.CellValue = new CellValue();
            }

            if (string.IsNullOrEmpty(displayValue))
            {
                cell.CellValue.Text = string.Empty;
                cell.DataType = CellValues.String;
                return;
            }

            if (double.TryParse(displayValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
            {
                cell.CellValue.Text = number.ToString(CultureInfo.InvariantCulture);
                cell.DataType = null;
                return;
            }

            cell.CellValue.Text = displayValue;
            cell.DataType = CellValues.String;
        }

        private void DeleteColumns(Worksheet worksheet, HashSet<int> columnsToDelete, CancellationToken cancellationToken)
        {
            if (worksheet == null || columnsToDelete == null || columnsToDelete.Count == 0) return;

            var sortedDelete = columnsToDelete.OrderBy(c => c).ToList();
            var deleteSet = new HashSet<int>(sortedDelete);

            UpdateColumnsDefinitionForDeletion(worksheet, sortedDelete);
            RemoveMergedCellsByColumn(worksheet, sortedDelete);

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return;

            foreach (var row in sheetData.Elements<Row>())
            {
                cancellationToken.ThrowIfCancellationRequested();

                var rowIndex = row.RowIndex?.Value ?? 0U;
                var cells = row.Elements<Cell>().ToList();
                var remainingCells = new List<Cell>();

                foreach (var cell in cells)
                {
                    var columnIndex = GetColumnIndex(cell.CellReference);
                    if (!columnIndex.HasValue) continue;

                    if (deleteSet.Contains(columnIndex.Value))
                    {
                        continue;
                    }

                    var shift = CountDeletedBefore(sortedDelete, columnIndex.Value);
                    if (shift > 0)
                    {
                        UpdateCellReference(cell, columnIndex.Value - shift, rowIndex);
                    }

                    remainingCells.Add(cell);
                }

                if (remainingCells.Count != cells.Count)
                {
                    row.RemoveAllChildren<Cell>();
                    foreach (var cell in remainingCells.OrderBy(c => GetColumnIndex(c.CellReference) ?? int.MaxValue))
                    {
                        row.AppendChild(cell);
                    }
                }
            }

            worksheet.Save();
        }

        private void UpdateColumnsDefinitionForDeletion(Worksheet worksheet, List<int> sortedDelete)
        {
            var columnsElement = worksheet.Elements<Columns>().FirstOrDefault();
            if (columnsElement == null) return;

            var columnList = columnsElement.Elements<Column>().ToList();

            foreach (var deleteIndex in sortedDelete)
            {
                var updated = new List<Column>();

                foreach (var column in columnList)
                {
                    var min = column.Min != null ? (int)column.Min.Value : 0;
                    var max = column.Max != null ? (int)column.Max.Value : 0;

                    if (max == 0 || min == 0)
                    {
                        updated.Add(column);
                        continue;
                    }

                    if (max < deleteIndex)
                    {
                        updated.Add(column);
                        continue;
                    }

                    if (min > deleteIndex)
                    {
                        column.Min = (uint)(min - 1);
                        column.Max = (uint)(max - 1);
                        updated.Add(column);
                        continue;
                    }

                    if (min == deleteIndex && max == deleteIndex)
                    {
                        continue;
                    }

                    if (min == deleteIndex)
                    {
                        column.Min = (uint)(deleteIndex + 1);
                        updated.Add(column);
                        continue;
                    }

                    if (max == deleteIndex)
                    {
                        column.Max = (uint)(deleteIndex - 1);
                        updated.Add(column);
                        continue;
                    }

                    var left = column.CloneNode(true) as Column;
                    if (left != null)
                    {
                        left.Min = column.Min;
                        left.Max = (uint)(deleteIndex - 1);
                        updated.Add(left);
                    }

                    column.Min = (uint)(deleteIndex + 1);
                    updated.Add(column);
                }

                columnList = updated;
            }

            columnsElement.RemoveAllChildren<Column>();
            foreach (var column in columnList.OrderBy(c => c.Min?.Value ?? 0U))
            {
                columnsElement.AppendChild(column);
            }
        }

        private void DeleteRows(Worksheet worksheet, HashSet<int> rowsToDelete, CancellationToken cancellationToken)
        {
            if (worksheet == null || rowsToDelete == null || rowsToDelete.Count == 0) return;

            var sortedDelete = rowsToDelete.OrderBy(r => r).ToList();
            RemoveMergedCellsByRow(worksheet, sortedDelete);

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return;

            var rows = sheetData.Elements<Row>().ToList();
            var remainingRows = new List<Row>();

            foreach (var row in rows)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var rowIndex = row.RowIndex != null ? (int)row.RowIndex.Value : 0;
                if (rowIndex == 0) continue;

                if (rowsToDelete.Contains(rowIndex))
                {
                    continue;
                }

                var shift = CountDeletedBefore(sortedDelete, rowIndex);
                if (shift > 0)
                {
                    var newRowIndex = (uint)(rowIndex - shift);
                    row.RowIndex = newRowIndex;

                    foreach (var cell in row.Elements<Cell>())
                    {
                        var columnIndex = GetColumnIndex(cell.CellReference);
                        if (columnIndex.HasValue)
                        {
                            UpdateCellReference(cell, columnIndex.Value, newRowIndex);
                        }
                    }
                }

                remainingRows.Add(row);
            }

            sheetData.RemoveAllChildren<Row>();
            foreach (var row in remainingRows.OrderBy(r => r.RowIndex?.Value ?? 0U))
            {
                sheetData.AppendChild(row);
            }

            worksheet.Save();
        }

        private void RemoveMergedCellsByColumn(Worksheet worksheet, List<int> sortedDelete)
        {
            var mergeCells = worksheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells == null) return;

            bool removed = false;
            foreach (var mergeCell in mergeCells.Elements<MergeCell>().ToList())
            {
                if (TryParseCellRange(mergeCell.Reference?.Value, out var startColumn, out var _, out var endColumn, out var _))
                {
                    if (sortedDelete.Any(c => c >= startColumn && c <= endColumn))
                    {
                        mergeCell.Remove();
                        removed = true;
                    }
                }
            }

            if (removed && !mergeCells.Elements<MergeCell>().Any())
            {
                mergeCells.Remove();
            }
        }

        private void RemoveMergedCellsByRow(Worksheet worksheet, List<int> sortedDelete)
        {
            var mergeCells = worksheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells == null) return;

            bool removed = false;
            foreach (var mergeCell in mergeCells.Elements<MergeCell>().ToList())
            {
                if (TryParseCellRange(mergeCell.Reference?.Value, out var _, out var startRow, out var _, out var endRow))
                {
                    if (sortedDelete.Any(r => r >= startRow && r <= endRow))
                    {
                        mergeCell.Remove();
                        removed = true;
                    }
                }
            }

            if (removed && !mergeCells.Elements<MergeCell>().Any())
            {
                mergeCells.Remove();
            }
        }

        private int? GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) return null;

            int index = 0;
            foreach (var ch in cellReference)
            {
                if (char.IsLetter(ch))
                {
                    index = index * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
                }
                else
                {
                    break;
                }
            }

            return index == 0 ? (int?)null : index;
        }

        private string GetColumnName(int index)
        {
            if (index <= 0) return string.Empty;

            var name = string.Empty;
            int current = index;

            while (current > 0)
            {
                current--;
                name = (char)('A' + current % 26) + name;
                current /= 26;
            }

            return name;
        }

        private void UpdateCellReference(Cell cell, int columnIndex, uint rowIndex)
        {
            cell.CellReference = $"{GetColumnName(columnIndex)}{rowIndex}";
        }

        private int CountDeletedBefore(IList<int> sortedDelete, int index)
        {
            int left = 0;
            int right = sortedDelete.Count;

            while (left < right)
            {
                int mid = (left + right) / 2;
                if (sortedDelete[mid] < index)
                {
                    left = mid + 1;
                }
                else
                {
                    right = mid;
                }
            }

            return left;
        }

        private bool TryParseCellRange(string range, out int startColumn, out int startRow, out int endColumn, out int endRow)
        {
            startColumn = startRow = endColumn = endRow = 0;
            if (string.IsNullOrEmpty(range)) return false;

            var parts = range.Split(':');
            if (parts.Length != 2) return false;

            return TryParseCellReference(parts[0], out startColumn, out startRow) &&
                   TryParseCellReference(parts[1], out endColumn, out endRow);
        }

        private bool TryParseCellReference(string reference, out int columnIndex, out int rowIndex)
        {
            columnIndex = 0;
            rowIndex = 0;

            if (string.IsNullOrEmpty(reference)) return false;

            int i = 0;
            while (i < reference.Length && char.IsLetter(reference[i]))
            {
                columnIndex = columnIndex * 26 + (char.ToUpperInvariant(reference[i]) - 'A' + 1);
                i++;
            }

            if (columnIndex == 0 || i >= reference.Length) return false;

            var rowPart = reference.Substring(i);
            if (!int.TryParse(rowPart, NumberStyles.None, CultureInfo.InvariantCulture, out rowIndex))
            {
                return false;
            }

            return rowIndex > 0;
        }

        private Sheet GetSheet(WorkbookPart workbookPart, WorksheetPart worksheetPart)
        {
            var sheetId = workbookPart.GetIdOfPart(worksheetPart);
            return workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Id?.Value == sheetId);
        }
    }

    public class OtherOperationOptions
    {
        public string FolderPath { get; set; } = string.Empty;
        public bool IncludeSubDirectories { get; set; }
        public double IgnoreFileSizeMb { get; set; }
        public bool IncludeInvisibleSheets { get; set; }
        public int MaxParallelism { get; set; }
        public bool RemoveAllShapes { get; set; }
        public bool RemoveAllFormulas { get; set; }
        public bool EnableLogOutput { get; set; }
        public bool DeleteByKeywordEnabled { get; set; }
        public bool DeleteByKeywordDirectionIsRow { get; set; }
        public bool DeleteByKeywordTargetAll { get; set; }
        public HashSet<int> DeleteByKeywordTargets { get; set; } = new HashSet<int>();
        public string[] DeleteByKeywordKeywords { get; set; } = Array.Empty<string>();
        public bool DeleteByKeywordFullMatch { get; set; }
        public bool DeleteByKeywordCaseSensitive { get; set; }
        public bool DeleteByKeywordWidthSensitive { get; set; }
    }

    public class OtherOperationResult
    {
        public int TotalFiles { get; set; }
        public int ProcessedFiles { get; set; }
        public int ModifiedFiles { get; set; }
        public int ErrorCount { get; set; }
    }
}
