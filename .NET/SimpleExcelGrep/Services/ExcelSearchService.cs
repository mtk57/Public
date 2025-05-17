using System;
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
        // メモリクリーンアップ間隔を増やして頻度を下げる
        private const int MEMORY_CLEANUP_INTERVAL = 200; // 以前は50ファイルごと、これを200ファイルごとに変更

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
        /// 指定されたフォルダ内のExcelファイルを検索（パフォーマンス改善版）
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
            ConcurrentQueue<SearchResult> resultQueue,
            Action<string> statusUpdateCallback,
            CancellationToken cancellationToken)
        {
            _logger.LogMessage($"SearchExcelFilesAsync 開始: フォルダ={folderPath}");
            _logger.LogMessage($"検索開始: キーワード='{keyword}', 正規表現={useRegex}, 最初のヒットのみ={firstHitOnly}, 図形内検索={searchShapes}, 無視ファイルサイズ(MB)={ignoreFileSizeMB}");

            List<SearchResult> results = new List<SearchResult>();
            ConcurrentBag<SearchResult> concurrentResults = new ConcurrentBag<SearchResult>();

            try
            {
                string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.AllDirectories)
                                          .Concat(Directory.GetFiles(folderPath, "*.xlsm", SearchOption.AllDirectories))
                                          .ToArray();
                _logger.LogMessage($"{excelFiles.Length} 個のExcelファイル(.xlsx, .xlsm)が見つかりました");

                int totalFiles = excelFiles.Length;
                int processedFiles = 0;
                int processedBatch = 0;

                var parallelOptions = new ParallelOptions
                {
                    MaxDegreeOfParallelism = maxParallelism,
                    CancellationToken = cancellationToken
                };

                await Task.Run(() =>
                {
                    // バッチ処理をやめて元のように全ファイルを一度に並列処理
                    Parallel.ForEach(excelFiles, parallelOptions, (filePath) =>
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        // 無視キーワードチェック
                        if (ignoreKeywords.Any(k => filePath.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            _logger.LogMessage($"無視キーワードが含まれるため処理をスキップ: {filePath}");
                            Interlocked.Increment(ref processedFiles);
                            statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル ({concurrentResults.Count} 件見つかりました)");
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
                                    statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル ({concurrentResults.Count} 件見つかりました)");
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
                            // 深いログ出力をスキップして処理速度を向上
                            if (processedFiles % 100 == 0)
                            {
                                _logger.LogMessage($"ファイル処理中: {filePath}");
                            }
                    
                            string extension = Path.GetExtension(filePath).ToLowerInvariant();

                            if (extension == ".xlsx" || extension == ".xlsm")
                            {
                                List<SearchResult> fileResults = SearchInXlsxFile(
                                    filePath, keyword, useRegex, regex, resultQueue,
                                    firstHitOnly, searchShapes, cancellationToken);

                                if (fileResults.Count > 0)
                                {
                                    _logger.LogMessage($"一致を発見: {filePath}, 結果: {fileResults.Count}件");
                                }

                                foreach (var result in fileResults)
                                {
                                    concurrentResults.Add(result);
                                }
                            }

                            int filesProcessed = Interlocked.Increment(ref processedFiles);
                    
                            // ステータス更新頻度を下げてオーバーヘッドを削減
                            if (filesProcessed % 10 == 0 || filesProcessed == totalFiles)
                            {
                                statusUpdateCallback($"処理中... {filesProcessed}/{totalFiles} ファイル ({concurrentResults.Count} 件見つかりました)");
                            }
                    
                            // GCの頻度を下げる
                            if (Interlocked.Increment(ref processedBatch) % MEMORY_CLEANUP_INTERVAL == 0)
                            {
                                // GCを軽量化（詳細ログも削除）
                                GC.Collect(1, GCCollectionMode.Optimized, false);
                            }
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
                            statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル ({concurrentResults.Count} 件見つかりました)");
                        }
                    });
            
                    // 並列処理完了後にConcurrentBagから普通のリストに移す
                    results.AddRange(concurrentResults);
                }, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                _logger.LogMessage("SearchExcelFilesAsync内のタスクがキャンセルされました。");
                results.AddRange(concurrentResults);
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"SearchExcelFilesAsync内のタスクで予期せぬエラー: {ex.Message}");
                results.AddRange(concurrentResults);
                throw;
            }
            finally
            {
                // 終了時には完全なメモリクリーンアップを実行
                _logger.LogMessage("SearchExcelFilesAsync 終了時のメモリクリーンアップ実行");
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

            _logger.LogMessage($"SearchExcelFilesAsync 完了: {results.Count}件の結果");
            return results;
        }

         /// <summary>
        /// 単一のXLSX/XLSMファイルを検索（最適化版）
        /// </summary>
        private List<SearchResult> SearchInXlsxFile(
            string filePath,
            string keyword,
            bool useRegex,
            Regex regex,
            ConcurrentQueue<SearchResult> pendingResults,
            bool firstHitOnly,
            bool searchShapes,
            CancellationToken cancellationToken)
        {
            List<SearchResult> localResults = new List<SearchResult>();
            bool foundHitInFile = false;

            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    if (workbookPart == null) return localResults;

                    SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    SharedStringTable sharedStringTable = sharedStringTablePart?.SharedStringTable;

                    foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (firstHitOnly && foundHitInFile) break;

                        string sheetName = GetSheetName(workbookPart, worksheetPart) ?? "不明なシート";

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
                
                        // 重いクリーンアップを避け、軽量なクリーンアップだけ行う
                        // ガベージコレクションに任せる方が速い場合が多い
                        LightCleanupWorksheetPart(worksheetPart);
                    }
            
                    // SharedStringTableの最小限のクリーンアップ
                    if (sharedStringTable != null && sharedStringTable.HasChildren)
                    {
                        // 完全クリーンアップを避け、参照解除だけにする
                        sharedStringTable = null;
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

        /// <summary>
        /// WorksheetPartのリソースを解放
        /// </summary>
        private void CleanupWorksheetPart(WorksheetPart worksheetPart)
        {
            if (worksheetPart == null) return;

            try
            {
                // シートデータのクリーンアップ
                if (worksheetPart.Worksheet != null)
                {
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null && sheetData.HasChildren)
                    {
                        foreach (var row in sheetData.Elements<Row>())
                        {
                            if (row.HasChildren)
                            {
                                foreach (var cell in row.Elements<Cell>())
                                {
                                    if (cell.CellValue != null)
                                    {
                                        cell.CellValue.RemoveAllChildren();
                                    }
                                    cell.RemoveAllChildren();
                                }
                                row.RemoveAllChildren();
                            }
                        }
                        sheetData.RemoveAllChildren();
                    }
            
                    // WorksheetのOtherChildren (Hyperlinks, MergedCells, etc.)
                    foreach (var element in worksheetPart.Worksheet.ChildElements)
                    {
                        if (element.HasChildren)
                        {
                            element.RemoveAllChildren();
                        }
                    }
                }
        
                // 図形データのクリーンアップ
                if (worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null)
                {
                    var drawing = worksheetPart.DrawingsPart.WorksheetDrawing;
            
                    // TwoCellAnchorの子要素をクリーンアップ
                    foreach (var anchor in drawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
                    {
                        if (anchor.HasChildren)
                        {
                            var shapes = anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().ToList();
                            foreach (var shape in shapes)
                            {
                                if (shape.TextBody != null && shape.TextBody.HasChildren)
                                {
                                    shape.TextBody.RemoveAllChildren();
                                }
                                shape.RemoveAllChildren();
                            }
                    
                            var graphicFrames = anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>().ToList();
                            foreach (var frame in graphicFrames)
                            {
                                if (frame.Graphic != null && frame.Graphic.GraphicData != null)
                                {
                                    frame.Graphic.GraphicData.RemoveAllChildren();
                                }
                                if (frame.Graphic != null)
                                {
                                    frame.Graphic.RemoveAllChildren();
                                }
                                frame.RemoveAllChildren();
                            }
                    
                            anchor.RemoveAllChildren();
                        }
                    }
            
                    drawing.RemoveAllChildren();
                }
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"WorksheetPartクリーンアップエラー: {ex.Message}");
            }
        }

        /// <summary>
        /// WorkbookPartのリソースを解放
        /// </summary>
        private void CleanupWorkbookPart(WorkbookPart workbookPart)
        {
            if (workbookPart == null) return;

            try
            {
                // Sheets, DefinedNames, BookViewsのクリーンアップ
                if (workbookPart.Workbook != null)
                {
                    var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                    if (sheets != null && sheets.HasChildren)
                    {
                        sheets.RemoveAllChildren();
                    }
            
                    var definedNames = workbookPart.Workbook.GetFirstChild<DefinedNames>();
                    if (definedNames != null && definedNames.HasChildren)
                    {
                        definedNames.RemoveAllChildren();
                    }
            
                    var bookViews = workbookPart.Workbook.GetFirstChild<BookViews>();
                    if (bookViews != null && bookViews.HasChildren)
                    {
                        bookViews.RemoveAllChildren();
                    }
                }
        
                // Styles, Theme, Calculation Propertiesなどの追加部品もクリーンアップ
                if (workbookPart.WorkbookStylesPart != null)
                {
                    var styles = workbookPart.WorkbookStylesPart.Stylesheet;
                    if (styles != null)
                    {
                        // スタイル情報をクリア
                        styles.RemoveAllChildren();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"WorkbookPartクリーンアップエラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 最適なバッチサイズを計算
        /// </summary>
        private int CalculateOptimalBatchSize(string[] files, double ignoreFileSizeMB, int maxParallelism)
        {
            // ファイル数が少ない場合はバッチ分けしない
            if (files.Length <= maxParallelism * 2)
            {
                return files.Length;
            }

            try
            {
                // サンプルファイルのサイズを取得してバッチサイズを推定
                long totalSampleSize = 0;
                int sampleCount = Math.Min(10, files.Length);
        
                for (int i = 0; i < sampleCount; i++)
                {
                    try
                    {
                        FileInfo fileInfo = new FileInfo(files[i]);
                        totalSampleSize += fileInfo.Length;
                    }
                    catch
                    {
                        // ファイルアクセスエラーは無視
                    }
                }
        
                // 平均ファイルサイズ (バイト)
                double avgFileSizeMB = (double)totalSampleSize / (sampleCount * 1024 * 1024);
        
                // 使用可能メモリを考慮したバッチサイズ (目安: 1ファイルあたり平均サイズの5倍のメモリを使用すると仮定)
                long availableMemory = GC.GetTotalMemory(false);
                double memoryPerFileMB = Math.Max(1, avgFileSizeMB * 5); // 最低1MBと仮定
        
                // 利用可能メモリの80%を使用すると仮定
                double usableMemoryMB = Math.Max(100, availableMemory / (1024 * 1024) * 0.8);
        
                // バッチサイズを計算 (メモリベース)
                int memoryBasedBatchSize = (int)Math.Max(1, Math.Min(100, usableMemoryMB / memoryPerFileMB));
        
                // 並列処理数ベースのバッチサイズ (通常は並列数の2-5倍がバランスが良い)
                int threadBasedBatchSize = maxParallelism * 3;
        
                // 最終的なバッチサイズを決定（小さい方を選択）
                int batchSize = Math.Min(memoryBasedBatchSize, threadBasedBatchSize);
        
                // 極端に小さいか大きい値にならないよう制限
                return Math.Max(5, Math.Min(50, batchSize));
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"バッチサイズ計算エラー: {ex.Message}");
                return maxParallelism * 2; // デフォルト値
            }
        }

        /// <summary>
        /// WorksheetPartの軽量クリーンアップ
        /// </summary>
        private void LightCleanupWorksheetPart(WorksheetPart worksheetPart)
        {
            // 重要なリソースの解放のみに留める
            // ほとんどの場合は明示的な解放は不要（ガベージコレクション任せが速い）
    
            // 大量のメモリを消費している可能性がある図形データのみクリーンアップ
            try
            {
                if (worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null)
                {
                    // 参照を解除するだけ
                    worksheetPart.DrawingsPart.WorksheetDrawing = null;
                }
            }
            catch
            {
                // エラーは無視（速度優先）
            }
        }
    }
}