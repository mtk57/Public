using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
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
        /// 指定されたフォルダ内のExcelファイルを検索（バランス改善版）
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

            List<SearchResult> results = new List<SearchResult>();
            // スレッドセーフなリストでなく、定期的に同期する方式
            List<List<SearchResult>> threadLocalResults = new List<List<SearchResult>>();
            for (int i = 0; i < maxParallelism; i++)
            {
                threadLocalResults.Add(new List<SearchResult>());
            }

            try
            {
                string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.AllDirectories)
                                          .Concat(Directory.GetFiles(folderPath, "*.xlsm", SearchOption.AllDirectories))
                                          .ToArray();
                _logger.LogMessage($"{excelFiles.Length} 個のExcelファイル(.xlsx, .xlsm)が見つかりました");

                int totalFiles = excelFiles.Length;
                int processedFiles = 0;
                int lastReportedCount = 0;

                // 一定間隔でメインリストに結果を同期するためのタイマー
                using (var syncTimer = new System.Threading.Timer(state => 
                {
                    try 
                    {
                        SynchronizeResults(threadLocalResults, results, resultQueue);
                        // ファイルの20%処理するごとにGCを強制実行
                        if (processedFiles - lastReportedCount > totalFiles * 0.2)
                        {
                            lastReportedCount = processedFiles;
                            // 強力なGCを実行
                            _logger.LogMessage($"進捗20%毎のメモリクリーンアップを実行 ({processedFiles}/{totalFiles})");
                            GC.Collect(2, GCCollectionMode.Forced, true, true);
                            GC.WaitForPendingFinalizers();
                            GC.Collect(2, GCCollectionMode.Forced, true, true);
                        }
                    } 
                    catch 
                    {
                        // タイマーでの例外は無視
                    }
                }, null, 2000, 2000))
                {
                    var parallelOptions = new ParallelOptions
                    {
                        MaxDegreeOfParallelism = maxParallelism,
                        CancellationToken = cancellationToken
                    };

                    await Task.Run(() =>
                    {
                        Parallel.ForEach(excelFiles, parallelOptions, (filePath, loopState, index) =>
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            int threadIndex = (int)index % maxParallelism;

                            // 無視キーワードチェック
                            if (ignoreKeywords.Any(k => filePath.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                            {
                                Interlocked.Increment(ref processedFiles);
                                if (processedFiles % 50 == 0 || processedFiles == totalFiles)
                                {
                                    statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル (結果: {results.Count + threadLocalResults.Sum(list => list.Count)})");
                                }
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
                                        Interlocked.Increment(ref processedFiles);
                                        if (processedFiles % 50 == 0 || processedFiles == totalFiles)
                                        {
                                            statusUpdateCallback($"処理中... {processedFiles}/{totalFiles} ファイル (結果: {results.Count + threadLocalResults.Sum(list => list.Count)})");
                                        }
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
                                string extension = Path.GetExtension(filePath).ToLowerInvariant();

                                if (extension == ".xlsx" || extension == ".xlsm")
                                {
                                    List<SearchResult> fileResults = SearchInXlsxFile(
                                        filePath, keyword, useRegex, regex, resultQueue,
                                        firstHitOnly, searchShapes, cancellationToken);

                                    // スレッドローカルの結果リストに追加（スレッドセーフ）
                                    if (fileResults.Count > 0)
                                    {
                                        lock (threadLocalResults[threadIndex])
                                        {
                                            threadLocalResults[threadIndex].AddRange(fileResults);
                                        }
                                
                                        // リアルタイム表示モード時のみキューに追加
                                        if (isRealTimeDisplay)
                                        {
                                            foreach (var result in fileResults)
                                            {
                                                resultQueue.Enqueue(result);
                                            }
                                        }
                                    }
                                }

                                int filesProcessed = Interlocked.Increment(ref processedFiles);
                                if (filesProcessed % 50 == 0 || filesProcessed == totalFiles)
                                {
                                    statusUpdateCallback($"処理中... {filesProcessed}/{totalFiles} ファイル (結果: {results.Count + threadLocalResults.Sum(list => list.Count)})");
                                }
                        
                                // 100ファイル毎に軽量GCを実行
                                if (filesProcessed % 100 == 0)
                                {
                                    GC.Collect(0, GCCollectionMode.Optimized, false);
                                }
                            }
                            catch (OperationCanceledException)
                            {
                                _logger.LogMessage($"ファイル処理がキャンセルされました: {filePath}");
                                throw;
                            }
                            catch (Exception ex)
                            {
                                _logger.LogMessage($"ファイル処理エラー: {filePath}, {ex.Message}");
                                Interlocked.Increment(ref processedFiles);
                            }
                        });
                
                        // 最終的に全ての結果を同期
                        SynchronizeResults(threadLocalResults, results, resultQueue);
                    }, cancellationToken);
                }
            }
            catch (OperationCanceledException)
            {
                _logger.LogMessage("SearchExcelFilesAsync内のタスクがキャンセルされました。");
                // キャンセル時も結果を同期
                SynchronizeResults(threadLocalResults, results, resultQueue);
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"SearchExcelFilesAsync内のタスクで予期せぬエラー: {ex.Message}");
                SynchronizeResults(threadLocalResults, results, resultQueue);
                throw;
            }
            finally
            {
                // 終了時には完全なメモリクリーンアップを実行
                _logger.LogMessage("SearchExcelFilesAsync 終了時のメモリクリーンアップ実行");
                GC.Collect(2, GCCollectionMode.Forced, true, true);
                GC.WaitForPendingFinalizers();
                GC.Collect(2, GCCollectionMode.Forced, true, true);
                _logger.LogMessage($"現在のメモリ使用量: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)}MB");
            }

            _logger.LogMessage($"SearchExcelFilesAsync 完了: {results.Count}件の結果");
            return results;
        }

         /// <summary>
        /// 単一のXLSX/XLSMファイルを検索（バランス改善版）
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
                
                        // バランスのとれたクリーンアップを実行
                        BalancedCleanupWorksheetPart(worksheetPart);
                    }
            
                    // シェアードストリングテーブルは明示的にクリーンアップ
                    if (sharedStringTable != null && sharedStringTable.HasChildren)
                    {
                        // トップレベルの子要素だけを削除（深すぎるクリーンアップは避ける）
                        sharedStringTable.RemoveAllChildren();
                    }
            
                    // ワークブックパーツのバランスのとれたクリーンアップ
                    BalancedCleanupWorkbookPart(workbookPart);
                }
            }
            catch (OperationCanceledException)
            {
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

        /// <summary>
        /// WorksheetPartのバランスのとれたクリーンアップ
        /// </summary>
        private void BalancedCleanupWorksheetPart(WorksheetPart worksheetPart)
        {
            if (worksheetPart == null) return;

            try
            {
                // 最も重要な図形データ部分のクリーンアップ
                if (worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null)
                {
                    var drawing = worksheetPart.DrawingsPart.WorksheetDrawing;
            
                    // TwoCellAnchorの子要素のうち、大きなデータを持つものだけクリーンアップ
                    foreach (var anchor in drawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
                    {
                        if (anchor.HasChildren)
                        {
                            // 特にテキストを持つShapeやGraphicFrameは確実にクリーンアップ
                            var shapes = anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().ToList();
                            foreach (var shape in shapes)
                            {
                                if (shape.TextBody != null && shape.TextBody.HasChildren)
                                {
                                    shape.TextBody.RemoveAllChildren();
                                }
                            }
                    
                            var graphicFrames = anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>().ToList();
                            foreach (var frame in graphicFrames)
                            {
                                if (frame.Graphic != null && frame.Graphic.GraphicData != null)
                                {
                                    frame.Graphic.GraphicData.RemoveAllChildren();
                                }
                            }
                        }
                    }
                }
        
                // シートデータは条件付きでクリーンアップ
                if (worksheetPart.Worksheet != null)
                {
                    // セルの値だけをクリア（構造は維持）
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null && sheetData.HasChildren)
                    {
                        foreach (var row in sheetData.Elements<Row>())
                        {
                            foreach (var cell in row.Elements<Cell>())
                            {
                                if (cell.CellValue != null)
                                {
                                    // セル値のテキスト内容だけクリア
                                    cell.CellValue.Text = string.Empty;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // エラーをログに記録するが、処理は継続
                _logger.LogMessage($"WorksheetPartクリーンアップエラー: {ex.Message}");
            }
        }

        /// <summary>
        /// WorkbookPartのバランスのとれたクリーンアップ
        /// </summary>
        private void BalancedCleanupWorkbookPart(WorkbookPart workbookPart)
        {
            // 最低限のクリーンアップのみ実行
            if (workbookPart?.WorkbookStylesPart?.Stylesheet != null)
            {
                try
                {
                    // スタイル情報だけをクリア
                    workbookPart.WorkbookStylesPart.Stylesheet.CellStyles = null;
                    workbookPart.WorkbookStylesPart.Stylesheet.CellStyleFormats = null;
                }
                catch
                {
                    // エラーは無視
                }
            }
        }

        /// <summary>
        /// スレッドローカルの結果をメインリストに同期
        /// </summary>
        private void SynchronizeResults(List<List<SearchResult>> threadLocalResults, List<SearchResult> mainResults, ConcurrentQueue<SearchResult> resultQueue)
        {
            lock (mainResults)
            {
                for (int i = 0; i < threadLocalResults.Count; i++)
                {
                    lock (threadLocalResults[i])
                    {
                        // 新しい結果のみ追加
                        int startCount = mainResults.Count;
                        mainResults.AddRange(threadLocalResults[i]);
                
                        // すでに同期した結果はクリア
                        threadLocalResults[i].Clear();
                
                        _logger.LogMessage($"結果同期: スレッド {i} から {mainResults.Count - startCount} 件追加");
                    }
                }
            }
        }
    }
}