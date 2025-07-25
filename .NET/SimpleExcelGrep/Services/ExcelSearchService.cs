﻿using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SimpleExcelGrep.Models;

namespace SimpleExcelGrep.Services
{
    /// <summary>
    /// Excelファイル検索を担当するサービス
    /// </summary>
    internal class ExcelSearchService
    {
        private readonly LogService _logger;
        private readonly ShapeTextExtractor _shapeTextExtractor;

        /// <summary>
        /// ExcelSearchServiceのコンストラクタ
        /// </summary>
        /// <param name="logger">ログサービス</param>
        public ExcelSearchService(LogService logger)
        {
            _logger = logger;
            _shapeTextExtractor = new ShapeTextExtractor();
        }

        /// <summary>
        /// (GREP検索モード) 指定されたフォルダ内のExcelファイルを検索
        /// </summary>
        public async Task<List<SearchResult>> SearchExcelFilesAsync(
            string folderPath,
            string keyword,
            bool useRegex,
            Regex regex,
            List<string> ignoreKeywords,
            bool isRealTimeDisplay,
            bool searchShapes,
            bool firstHitOnly,
            int maxParallelism,
            double ignoreFileSizeMB,
            bool searchSubDirectories,
            bool searchInvisibleSheets,
            ConcurrentQueue<SearchResult> resultQueue,
            Action<string> statusUpdateCallback,
            CancellationToken cancellationToken)
        {
            _logger.LogMessage($"SearchExcelFilesAsync 開始: フォルダ={folderPath}");
            _logger.LogMessage($"GREP検索開始: キーワード='{keyword}', 正規表現={useRegex}, 最初のヒットのみ={firstHitOnly}, 図形内検索={searchShapes}, 無視ファイルサイズ(MB)={ignoreFileSizeMB}, サブフォルダ対象={searchSubDirectories}, 非表示シート対象={searchInvisibleSheets}");

            List<SearchResult> results = new List<SearchResult>();

            try
            {
                var searchOption = searchSubDirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", searchOption)
                                          .Concat(Directory.GetFiles(folderPath, "*.xlsm", searchOption))
                                          .ToArray();
                _logger.LogMessage($"{excelFiles.Length} 個のExcelファイル(.xlsx, .xlsm)が見つかりました");

                int totalFiles = excelFiles.Length;
                int processedFiles = 0;

                var parallelOptions = new ParallelOptions
                {
                    MaxDegreeOfParallelism = maxParallelism,
                    CancellationToken = cancellationToken
                };

                await Task.Run(() =>
                {
                    Parallel.ForEach(excelFiles, parallelOptions, (filePath) =>
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        // 無視キーワードチェック
                        if (ignoreKeywords.Any(k => filePath.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            _logger.LogMessage($"無視キーワードが含まれるため処理をスキップ: {filePath}");
                            Interlocked.Increment(ref processedFiles);
                            statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル ({results.Count} 件見つかりました)");
                            return;
                        }

                        // ファイルサイズチェック
                        if (ignoreFileSizeMB > 0)
                        {
                            try
                            {
                                FileInfo fileInfo = new FileInfo(filePath);
                                double fileSizeMB = (double)fileInfo.Length / (1024 * 1024);
                                if (fileSizeMB > ignoreFileSizeMB)
                                {
                                    _logger.LogMessage($"ファイルサイズ超過のためスキップ ({fileSizeMB:F2}MB > {ignoreFileSizeMB}MB): {filePath}");
                                    Interlocked.Increment(ref processedFiles);
                                    statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル ({results.Count} 件見つかりました)");
                                    return;
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.LogMessage($"ファイルサイズ取得エラー: {filePath}, {ex.Message}");
                            }
                        }

                        try
                        {
                            _logger.LogMessage($"ファイル処理開始: {filePath}");
                            List<SearchResult> fileResults = SearchInXlsxFile(
                                filePath, keyword, useRegex, regex, resultQueue,
                                firstHitOnly, searchShapes, searchInvisibleSheets, cancellationToken);

                            _logger.LogMessage($"ファイル処理完了: {filePath}, 見つかった結果(セル+図形): {fileResults.Count}件");

                            lock (results)
                            {
                                results.AddRange(fileResults);
                            }

                            Interlocked.Increment(ref processedFiles);
                            statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル ({results.Count} 件見つかりました)");
                        }
                        catch (OperationCanceledException)
                        {
                            _logger.LogMessage($"ファイル処理がキャンセルされました: {filePath}");
                            throw;
                        }
                        catch (Exception ex)
                        {
                            _logger.LogMessage($"ファイル処理エラー: {filePath}, {ex.GetType().Name}: {ex.Message}");
                            Interlocked.Increment(ref processedFiles);
                            statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル ({results.Count} 件見つかりました)");
                        }
                    });
                }, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                _logger.LogMessage("SearchExcelFilesAsync内のタスクがキャンセルされました。");
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"SearchExcelFilesAsync内のタスクで予期せぬエラー: {ex.Message}");
                throw;
            }

            _logger.LogMessage($"SearchExcelFilesAsync 完了: {results.Count}件の結果");
            return results;
        }

        /// <summary>
        /// (セル検索モード) 指定されたセルアドレスの値を検索
        /// </summary>
        public async Task<List<SearchResult>> SearchCellsByAddressAsync(
            string folderPath,
            string[] cellAddresses,
            List<string> ignoreKeywords,
            int maxParallelism,
            double ignoreFileSizeMB,
            bool searchSubDirectories,
            bool searchInvisibleSheets,
            ConcurrentQueue<SearchResult> resultQueue,
            Action<string> statusUpdateCallback,
            CancellationToken cancellationToken)
        {
            _logger.LogMessage($"SearchCellsByAddressAsync 開始: フォルダ={folderPath}");
            _logger.LogMessage($"セル検索開始: アドレス='{string.Join(", ", cellAddresses)}', 無視ファイルサイズ(MB)={ignoreFileSizeMB}, サブフォルダ対象={searchSubDirectories}, 非表示シート対象={searchInvisibleSheets}");

            List<SearchResult> results = new List<SearchResult>();
            
            try
            {
                var searchOption = searchSubDirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", searchOption)
                                          .Concat(Directory.GetFiles(folderPath, "*.xlsm", searchOption))
                                          .ToArray();
                _logger.LogMessage($"{excelFiles.Length} 個のExcelファイルが見つかりました");

                int totalFiles = excelFiles.Length;
                int processedFiles = 0;

                var parallelOptions = new ParallelOptions
                {
                    MaxDegreeOfParallelism = maxParallelism,
                    CancellationToken = cancellationToken
                };

                await Task.Run(() =>
                {
                    Parallel.ForEach(excelFiles, parallelOptions, (filePath) =>
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        // 無視キーワードチェック
                        if (ignoreKeywords.Any(k => filePath.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            _logger.LogMessage($"無視キーワードが含まれるためスキップ: {filePath}");
                            Interlocked.Increment(ref processedFiles);
                            statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル");
                            return;
                        }

                        // ファイルサイズチェック
                        if (ignoreFileSizeMB > 0)
                        {
                            try
                            {
                                FileInfo fileInfo = new FileInfo(filePath);
                                double fileSizeMB = (double)fileInfo.Length / (1024 * 1024);
                                if (fileSizeMB > ignoreFileSizeMB)
                                {
                                    _logger.LogMessage($"ファイルサイズ超過のためスキップ ({fileSizeMB:F2}MB > {ignoreFileSizeMB}MB): {filePath}");
                                    Interlocked.Increment(ref processedFiles);
                                    statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル");
                                    return;
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.LogMessage($"ファイルサイズ取得エラー: {filePath}, {ex.Message}");
                            }
                        }

                        try
                        {
                            _logger.LogMessage($"セル値取得開始: {filePath}");
                            List<SearchResult> fileResults = SearchCellsInXlsxFile(
                                filePath, cellAddresses, searchInvisibleSheets, resultQueue, cancellationToken);
                            
                            _logger.LogMessage($"セル値取得完了: {filePath}, {fileResults.Count}件");

                            lock (results)
                            {
                                results.AddRange(fileResults);
                            }

                            Interlocked.Increment(ref processedFiles);
                            statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル");
                        }
                        catch (OperationCanceledException)
                        {
                             _logger.LogMessage($"ファイル処理がキャンセルされました: {filePath}");
                            throw;
                        }
                        catch (Exception ex)
                        {
                             _logger.LogMessage($"ファイル処理エラー: {filePath}, {ex.GetType().Name}: {ex.Message}");
                            Interlocked.Increment(ref processedFiles);
                            statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル");
                        }
                    });
                }, cancellationToken);

            }
            catch (OperationCanceledException)
            {
                 _logger.LogMessage("SearchCellsByAddressAsync内のタスクがキャンセルされました。");
                throw;
            }

            return results;
        }

        /// <summary>
        /// (セル検索モード) 単一のXLSX/XLSMファイルから指定されたセルの値を取得
        /// </summary>
        private List<SearchResult> SearchCellsInXlsxFile(
            string filePath,
            string[] cellAddresses,
            bool searchInvisibleSheets,
            ConcurrentQueue<SearchResult> pendingResults,
            CancellationToken cancellationToken)
        {
            List<SearchResult> localResults = new List<SearchResult>();

            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    if (workbookPart == null || workbookPart.Workbook == null) return localResults;

                    SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    SharedStringTable sharedStringTable = sharedStringTablePart?.SharedStringTable;

                    foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        string sheetId = workbookPart.GetIdOfPart(worksheetPart);
                        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Id?.Value == sheetId);

                        if (!searchInvisibleSheets && sheet != null && sheet.State != null && sheet.State.Value != SheetStateValues.Visible)
                        {
                            _logger.LogMessage($"非表示シートのためスキップ: {sheet.Name?.Value} in {filePath}");
                            continue;
                        }
                        
                        string sheetName = sheet?.Name?.Value ?? "不明なシート";

                        foreach (string address in cellAddresses)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            Cell cell = worksheetPart.Worksheet.Descendants<Cell>()
                                .FirstOrDefault(c => c.CellReference?.Value == address);

                            string cellValue = (cell != null) ? GetCellValue(cell, sharedStringTable) : "(セルなし)";

                            SearchResult result = new SearchResult
                            {
                                FilePath = filePath,
                                SheetName = sheetName,
                                CellPosition = address,
                                CellValue = cellValue
                            };

                            pendingResults.Enqueue(result);
                            localResults.Add(result);

                            _logger.LogMessage($"セル値取得: {filePath} - {sheetName} - {address} - '{TruncateString(result.CellValue)}'");
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                _logger.LogMessage($"ファイル処理がキャンセルされました: {filePath}");
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"Excel処理エラー: {filePath}, {ex.GetType().Name}: {ex.Message}");
            }

            return localResults;
        }


        /// <summary>
        /// (GREP検索モード) 単一のXLSX/XLSMファイルを検索
        /// </summary>
        private List<SearchResult> SearchInXlsxFile(
            string filePath,
            string keyword,
            bool useRegex,
            Regex regex,
            ConcurrentQueue<SearchResult> pendingResults,
            bool firstHitOnly,
            bool searchShapes,
            bool searchInvisibleSheets,
            CancellationToken cancellationToken)
        {
            List<SearchResult> localResults = new List<SearchResult>();
            bool foundHitInFile = false;

            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    if (workbookPart == null || workbookPart.Workbook == null) return localResults;

                    SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    SharedStringTable sharedStringTable = sharedStringTablePart?.SharedStringTable;

                    foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        string sheetId = workbookPart.GetIdOfPart(worksheetPart);
                        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Id?.Value == sheetId);

                        if (!searchInvisibleSheets && sheet != null && sheet.State != null && sheet.State.Value != SheetStateValues.Visible)
                        {
                            _logger.LogMessage($"非表示シートのためスキップ: {sheet.Name?.Value} in {filePath}");
                            continue;
                        }
                        
                        if (firstHitOnly && foundHitInFile) break;
                        
                        string sheetName = sheet?.Name?.Value ?? "不明なシート";

                        // 1. セル内のテキスト検索
                        foundHitInFile = SearchInCells(
                            filePath, sheetName, worksheetPart, sharedStringTable,
                            keyword, useRegex, regex, pendingResults,
                            localResults, firstHitOnly, foundHitInFile, cancellationToken);

                        // 2. 図形内のテキスト検索
                        if (searchShapes)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (firstHitOnly && foundHitInFile) break;

                            foundHitInFile = SearchInShapes(
                                filePath, sheetName, worksheetPart,
                                keyword, useRegex, regex, pendingResults,
                                localResults, firstHitOnly, foundHitInFile, cancellationToken);
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                _logger.LogMessage($"ファイル処理がキャンセルされました: {filePath}");
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"Excel処理エラー: {filePath}, {ex.GetType().Name}: {ex.Message}");
            }

            return localResults;
        }

        // 以下、既存のプライベートメソッド (SearchInCells, SearchInShapes, GetSheetName, GetCellValue, GetCellReference, TruncateString) は変更なし
        // ... (省略) ...
        /// <summary>
        /// ワークシート内のセルを検索
        /// </summary>
        private bool SearchInCells(
            string filePath, string sheetName, WorksheetPart worksheetPart,
            SharedStringTable sharedStringTable, string keyword, bool useRegex,
            Regex regex, ConcurrentQueue<SearchResult> pendingResults,
            List<SearchResult> localResults, bool firstHitOnly,
            bool foundHitInFile, CancellationToken cancellationToken)
        {
            if (worksheetPart.Worksheet == null) return foundHitInFile;

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return foundHitInFile;

            foreach (Row row in sheetData.Elements<Row>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (firstHitOnly && foundHitInFile) break;

                foreach (Cell cell in row.Elements<Cell>())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (firstHitOnly && foundHitInFile) break;

                    string cellValue = GetCellValue(cell, sharedStringTable);
                    if (string.IsNullOrEmpty(cellValue)) continue;

                    bool isMatch = useRegex && regex != null ?
                                   regex.IsMatch(cellValue) :
                                   cellValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;

                    if (isMatch)
                    {
                        SearchResult result = new SearchResult
                        {
                            FilePath = filePath,
                            SheetName = sheetName,
                            CellPosition = GetCellReference(cell),
                            CellValue = cellValue
                        };

                        pendingResults.Enqueue(result);
                        localResults.Add(result);
                        foundHitInFile = true;

                        _logger.LogMessage($"セル内一致: {filePath} - {sheetName} - {result.CellPosition} - '{TruncateString(result.CellValue)}'");

                        if (firstHitOnly) break;
                    }
                }
            }

            return foundHitInFile;
        }

        /// <summary>
        /// ワークシート内の図形を検索
        /// </summary>
        private bool SearchInShapes(
            string filePath, string sheetName, WorksheetPart worksheetPart,
            string keyword, bool useRegex, Regex regex,
            ConcurrentQueue<SearchResult> pendingResults, List<SearchResult> localResults,
            bool firstHitOnly, bool foundHitInFile, CancellationToken cancellationToken)
        {
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null || drawingsPart.WorksheetDrawing == null) return foundHitInFile;

            foreach (var twoCellAnchor in drawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (firstHitOnly && foundHitInFile) break;

                // Shapeからテキスト検索
                var shape = twoCellAnchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().FirstOrDefault();
                if (shape != null && shape.TextBody != null)
                {
                    string shapeText = _shapeTextExtractor.GetTextFromShapeTextBody(shape.TextBody);
                    if (!string.IsNullOrEmpty(shapeText))
                    {
                        bool isMatch = useRegex && regex != null ?
                                      regex.IsMatch(shapeText) :
                                      shapeText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;

                        if (isMatch)
                        {
                            SearchResult result = new SearchResult
                            {
                                FilePath = filePath,
                                SheetName = sheetName,
                                CellPosition = "図形内",
                                CellValue = TruncateString(shapeText)
                            };

                            pendingResults.Enqueue(result);
                            localResults.Add(result);
                            foundHitInFile = true;

                            _logger.LogMessage($"図形内一致 (Shape): {filePath} - {sheetName} - '{result.CellValue}'");

                            if (firstHitOnly) break;
                        }
                    }
                }

                // GraphicFrameからテキスト検索
                var graphicFrame = twoCellAnchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>().FirstOrDefault();
                if (graphicFrame != null)
                {
                    string frameText = _shapeTextExtractor.GetTextFromGraphicFrame(graphicFrame);
                    if (!string.IsNullOrEmpty(frameText))
                    {
                        bool isMatch = useRegex && regex != null ?
                                      regex.IsMatch(frameText) :
                                      frameText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;

                        if (isMatch)
                        {
                            SearchResult result = new SearchResult
                            {
                                FilePath = filePath,
                                SheetName = sheetName,
                                CellPosition = "図形内 (GF)",
                                CellValue = TruncateString(frameText)
                            };

                            pendingResults.Enqueue(result);
                            localResults.Add(result);
                            foundHitInFile = true;

                            _logger.LogMessage($"図形内一致 (GraphicFrame): {filePath} - {sheetName} - '{result.CellValue}'");

                            if (firstHitOnly) break;
                        }
                    }
                }
            }

            return foundHitInFile;
        }

        /// <summary>
        /// ワークシート名を取得
        /// </summary>
        private string GetSheetName(WorkbookPart workbookPart, WorksheetPart worksheetPart)
        {
            string sheetId = workbookPart.GetIdOfPart(worksheetPart);
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Id?.Value == sheetId);
            return sheet?.Name?.Value;
        }

        /// <summary>
        /// セルの値を取得
        /// </summary>
        private string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell == null || cell.CellValue == null)
                return string.Empty;

            string cellValueStr = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && sharedStringTable != null)
            {
                if (int.TryParse(cellValueStr, out int ssid) && ssid >= 0 && ssid < sharedStringTable.ChildElements.Count)
                {
                    SharedStringItem ssi = sharedStringTable.ChildElements[ssid] as SharedStringItem;
                    if (ssi != null)
                    {
                        // Text 要素の値を連結する
                        return string.Concat(ssi.Elements<Text>().Select(t => t.Text));
                    }
                }
                return string.Empty; // 共有文字列が見つからない場合
            }
            return cellValueStr;
        }

        /// <summary>
        /// セル参照（例：A1）を取得
        /// </summary>
        private string GetCellReference(Cell cell)
        {
            return cell.CellReference?.Value ?? string.Empty;
        }

        /// <summary>
        /// 文字列を指定の長さに切り詰める
        /// </summary>
        private string TruncateString(string value, int maxLength = 255)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength) + "...";
        }
    }
}