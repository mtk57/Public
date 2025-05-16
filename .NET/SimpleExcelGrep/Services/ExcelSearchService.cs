using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics; // For Stopwatch and Process
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
        private DateTime _searchStartTime; // 検索開始時刻を保持

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
        /// 指定されたフォルダ内のExcelファイルを検索
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
            _searchStartTime = DateTime.Now; // 検索開始時刻を記録
            _logger.LogMessage($"SearchExcelFilesAsync 開始: フォルダ='{folderPath}', MaxParallelism={maxParallelism}");
            _logger.LogMessage($"検索パラメータ: キーワード='{keyword}', 正規表現={useRegex}, 最初のヒットのみ={firstHitOnly}, 図形内検索={searchShapes}, 無視ファイルサイズ(MB)={ignoreFileSizeMB}");

            List<SearchResult> results = new List<SearchResult>();
            long totalFilesProcessedForPerfLog = 0; // パフォーマンスログ用の処理済みファイルカウンター

            try
            {
                string[] excelFiles = Array.Empty<string>();
                try
                {
                    excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.AllDirectories)
                                              .Concat(Directory.GetFiles(folderPath, "*.xlsm", SearchOption.AllDirectories))
                                              .ToArray();
                    _logger.LogMessage($"{excelFiles.Length} 個のExcelファイル(.xlsx, .xlsm)が見つかりました。");
                }
                catch (Exception ex)
                {
                    _logger.LogMessage($"ファイル一覧の取得中にエラーが発生しました: {ex.Message}");
                    statusUpdateCallback($"エラー: ファイル一覧の取得に失敗しました。ログを確認してください。");
                    return results; // ファイル一覧取得エラー時は空の結果を返す
                }


                int totalFiles = excelFiles.Length;
                int processedFilesUICounter = 0; // UI表示用の処理済みファイルカウンター

                var parallelOptions = new ParallelOptions
                {
                    MaxDegreeOfParallelism = maxParallelism,
                    CancellationToken = cancellationToken
                };

                _logger.LogMessage($"Parallel.ForEach を開始します。MaxDegreeOfParallelism: {parallelOptions.MaxDegreeOfParallelism}");

                await Task.Run(() =>
                {
                    try
                    {
                        Parallel.ForEach(excelFiles, parallelOptions, (filePath, loopState) =>
                        {
                            // CancellationToken は Parallel.ForEach が内部で監視しているので、
                            // ループの最初で明示的にチェックする必要は必ずしもないが、より早く反応するためには有効。
                            if (cancellationToken.IsCancellationRequested)
                            {
                                _logger.LogMessage($"キャンセル要求を検知 (Parallel.ForEachループ内先頭): {filePath}");
                                loopState.Stop(); // 現在のイテレーションは完了させるが、新しいイテレーションは開始しない
                                return;
                            }

                            Stopwatch fileProcessingStopwatch = Stopwatch.StartNew();
                            double fileSizeMB = 0;
                            string displayFilePath = Path.GetFileName(filePath); // ログやUI表示用にファイル名だけにする

                            try
                            {
                                // 無視キーワードチェック
                                if (ignoreKeywords.Any(k => filePath.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                                {
                                    _logger.LogMessage($"無視キーワードが含まれるためスキップ: {displayFilePath}");
                                    Interlocked.Increment(ref processedFilesUICounter);
                                    statusUpdateCallback?.Invoke($"処理中... {processedFilesUICounter}/{totalFiles} ファイル ({results.Count} 件見つかりました)");
                                    fileProcessingStopwatch.Stop();
                                    long currentProcessedForPerf = Interlocked.Increment(ref totalFilesProcessedForPerfLog);
                                    TimeSpan elapsedSinceStart = DateTime.Now - _searchStartTime;
                                    double currentMemoryUsageMB = GetCurrentMemoryUsageMB();
                                    _logger.LogPerformanceData((int)currentProcessedForPerf, elapsedSinceStart.TotalSeconds, currentMemoryUsageMB, filePath, 0, 0);
                                    return;
                                }

                                // ファイルサイズチェック と ファイルサイズ取得
                                try
                                {
                                    FileInfo fileInfo = new FileInfo(filePath);
                                    fileSizeMB = (double)fileInfo.Length / (1024 * 1024);
                                    if (ignoreFileSizeMB > 0 && fileSizeMB > ignoreFileSizeMB)
                                    {
                                        _logger.LogMessage($"ファイルサイズ超過のためスキップ ({fileSizeMB:F2}MB > {ignoreFileSizeMB}MB): {displayFilePath}");
                                        Interlocked.Increment(ref processedFilesUICounter);
                                        statusUpdateCallback?.Invoke($"処理中... {processedFilesUICounter}/{totalFiles} ファイル ({results.Count} 件見つかりました)");
                                        fileProcessingStopwatch.Stop();
                                        long currentProcessedForPerf = Interlocked.Increment(ref totalFilesProcessedForPerfLog);
                                        TimeSpan elapsedSinceStart = DateTime.Now - _searchStartTime;
                                        double currentMemoryUsageMB = GetCurrentMemoryUsageMB();
                                        _logger.LogPerformanceData((int)currentProcessedForPerf, elapsedSinceStart.TotalSeconds, currentMemoryUsageMB, filePath, fileSizeMB, 0);
                                        return;
                                    }
                                }
                                catch (FileNotFoundException)
                                {
                                    _logger.LogMessage($"ファイルが見つかりません (サイズチェック時): {displayFilePath}。スキップします。");
                                    Interlocked.Increment(ref processedFilesUICounter);
                                    statusUpdateCallback?.Invoke($"処理中... {processedFilesUICounter}/{totalFiles} ファイル ({results.Count} 件見つかりました)");
                                    fileProcessingStopwatch.Stop();
                                    long currentProcessedForPerf = Interlocked.Increment(ref totalFilesProcessedForPerfLog);
                                    TimeSpan elapsedSinceStart = DateTime.Now - _searchStartTime;
                                    double currentMemoryUsageMB = GetCurrentMemoryUsageMB();
                                    _logger.LogPerformanceData((int)currentProcessedForPerf, elapsedSinceStart.TotalSeconds, currentMemoryUsageMB, filePath + " (NotFound_SizeCheck)", 0, 0);
                                    return;
                                }
                                catch (Exception ex)
                                {
                                    _logger.LogMessage($"ファイルサイズ取得エラー: {displayFilePath}, {ex.Message}。処理は続行します。");
                                    // エラーの場合はファイルサイズ0として処理を続行（スキップしない）
                                }


                                _logger.LogMessage($"ファイル処理開始: {displayFilePath}, サイズ: {fileSizeMB:F2} MB");
                                string extension = Path.GetExtension(filePath).ToLowerInvariant();

                                if (extension == ".xlsx" || extension == ".xlsm")
                                {
                                    List<SearchResult> fileResults = SearchInXlsxFile(
                                        filePath, keyword, useRegex, regex, resultQueue,
                                        firstHitOnly, searchShapes, cancellationToken);

                                    if (fileResults.Any())
                                    {
                                        lock (results) // resultsリストへのアクセスを同期
                                        {
                                            results.AddRange(fileResults);
                                        }
                                    }
                                    _logger.LogMessage($"ファイル処理完了: {displayFilePath}, 発見数: {fileResults.Count}, 所要時間: {fileProcessingStopwatch.ElapsedMilliseconds} ms");
                                }
                                else
                                {
                                    _logger.LogMessage($"サポート外の拡張子のためスキップ: {displayFilePath}");
                                }

                                Interlocked.Increment(ref processedFilesUICounter);
                                statusUpdateCallback?.Invoke($"処理中... {processedFilesUICounter}/{totalFiles} ファイル ({results.Count} 件見つかりました)");
                            }
                            catch (OperationCanceledException)
                            {
                                _logger.LogMessage($"ファイル処理がキャンセルされました (ループ内): {displayFilePath}");
                                fileProcessingStopwatch.Stop();
                                loopState.Stop(); // Parallel.ForEachの他のイテレーションを停止
                                // throw; // OperationCanceledException は Parallel.ForEach がキャッチして AggregateException に含める
                            }
                            catch (Exception ex) // 個別ファイル処理中の予期せぬエラー
                            {
                                _logger.LogMessage($"ファイル処理エラー (ループ内): {displayFilePath}, {ex.GetType().Name}: {ex.Message}. このファイルの処理をスキップします。");
                                _logger.LogMessage($"エラー詳細: {ex.StackTrace}");
                                Interlocked.Increment(ref processedFilesUICounter); // エラーでも処理済みファイルとしてカウント
                                statusUpdateCallback?.Invoke($"処理中... {processedFilesUICounter}/{totalFiles} ファイル ({results.Count} 件見つかりました)");
                            }
                            finally
                            {
                                fileProcessingStopwatch.Stop();
                                long currentProcessedCountForPerf = Interlocked.Increment(ref totalFilesProcessedForPerfLog);
                                TimeSpan elapsedTotal = DateTime.Now - _searchStartTime;
                                double memoryUsageMB = GetCurrentMemoryUsageMB();
                                _logger.LogPerformanceData((int)currentProcessedCountForPerf, elapsedTotal.TotalSeconds, memoryUsageMB, filePath, fileSizeMB, fileProcessingStopwatch.ElapsedMilliseconds);

                                // GC.Collect(); // 大量ファイル処理後、ここでGCを強制するのも一つの手だが、通常は推奨されない。パフォーマンスへの影響を慎重に評価。
                                // _logger.LogMessage($"GC.Collect() 実行後メモリ: {GetCurrentMemoryUsageMB():F2} MB for {displayFilePath}");
                            }
                        });
                    }
                    catch (OperationCanceledException) // Parallel.ForEach全体がキャンセルされた場合
                    {
                        _logger.LogMessage("Parallel.ForEach ループがキャンセルされました。");
                        throw; // 上位の呼び出し元に再スロー
                    }
                    catch (AggregateException ae) // Parallel.ForEach内で複数の例外が発生した場合
                    {
                        ae.Handle(ex =>
                        {
                            if (ex is OperationCanceledException)
                            {
                                _logger.LogMessage($"Parallel.ForEach 内で OperationCanceledException が発生しました。");
                                return true; // ハンドル済みとしてマーク
                            }
                            _logger.LogMessage($"Parallel.ForEach 内で予期せぬ例外: {ex.GetType().Name} - {ex.Message}");
                            _logger.LogMessage($"例外詳細: {ex.StackTrace}");
                            return true; // 他の例外もハンドル済みとしてマーク（処理は継続しないがクラッシュは防ぐ）
                        });
                        // AggregateException の内容によって再スローするか判断
                        if (ae.InnerExceptions.OfType<OperationCanceledException>().Any())
                        {
                            throw new OperationCanceledException("検索処理がキャンセルされました。", ae);
                        }
                        // その他の重要な例外があればここで再スローを検討
                        // throw new Exception("検索中に内部エラーが発生しました。", ae);
                    }


                }, cancellationToken); // Task.Run にも CancellationToken を渡す
            }
            catch (OperationCanceledException) // Task.Run や Parallel.ForEach からのキャンセル
            {
                _logger.LogMessage("SearchExcelFilesAsync 内のタスクがキャンセルされました。");
                throw; // MainForm 側で処理されるように再スロー
            }
            catch (Exception ex) // SearchExcelFilesAsyncレベルの予期せぬエラー
            {
                _logger.LogMessage($"SearchExcelFilesAsync 内で予期せぬエラー: {ex.GetType().Name} - {ex.Message}");
                _logger.LogMessage($"エラー詳細: {ex.StackTrace}");
                statusUpdateCallback?.Invoke($"重大なエラーが発生しました。ログを確認してください。");
                // throw; // 必要に応じて再スローしてアプリケーションを停止させるか、エラー状態として結果を返す
            }

            _logger.LogMessage($"SearchExcelFilesAsync 完了: {results.Count}件の結果。総処理時間: {(DateTime.Now - _searchStartTime).TotalSeconds:F2} 秒");
            return results;
        }

        /// <summary>
        /// 単一のXLSX/XLSMファイルを検索
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
            cancellationToken.ThrowIfCancellationRequested();
            List<SearchResult> localResults = new List<SearchResult>();
            bool foundHitInFile = false;
            string displayFilePath = Path.GetFileName(filePath);

            try
            {
                _logger.LogMessage($"SpreadsheetDocument.Open を試行: {displayFilePath}");
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    _logger.LogMessage($"SpreadsheetDocument.Open 成功: {displayFilePath}");
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    if (workbookPart == null)
                    {
                        _logger.LogMessage($"WorkbookPart が null です: {displayFilePath}。このファイルをスキップします。");
                        return localResults;
                    }

                    SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    SharedStringTable sharedStringTable = sharedStringTablePart?.SharedStringTable;
                    _logger.LogMessage($"SharedStringTable {(sharedStringTable != null ? "取得成功" : "取得失敗または存在せず")}: {displayFilePath}");


                    foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (firstHitOnly && foundHitInFile)
                        {
                            _logger.LogMessage($"最初のヒットが見つかったため、ファイル '{displayFilePath}' の残りのシート処理をスキップします。");
                            break;
                        }

                        string sheetName = GetSheetName(workbookPart, worksheetPart) ?? "不明なシート";
                        _logger.LogMessage($"シート処理開始: '{sheetName}' (ファイル: {displayFilePath})");

                        // 1. セル内のテキスト検索
                        if (!firstHitOnly || !foundHitInFile)
                        {
                            foundHitInFile = SearchInCells(
                                filePath, sheetName, worksheetPart, sharedStringTable,
                                keyword, useRegex, regex, pendingResults,
                                localResults, firstHitOnly, foundHitInFile, cancellationToken) || foundHitInFile; // OR代入
                        }


                        // 2. 図形内のテキスト検索
                        if (searchShapes && (!firstHitOnly || !foundHitInFile))
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                             _logger.LogMessage($"図形内検索開始: シート '{sheetName}' (ファイル: {displayFilePath})");
                            foundHitInFile = SearchInShapes(
                                filePath, sheetName, worksheetPart,
                                keyword, useRegex, regex, pendingResults,
                                localResults, firstHitOnly, foundHitInFile, cancellationToken) || foundHitInFile; // OR代入
                        }
                        _logger.LogMessage($"シート処理完了: '{sheetName}' (ファイル: {displayFilePath}), このシートでのヒット有無: {foundHitInFile}");
                    }
                } // usingブロック終了時にspreadsheetDocument.Dispose()が呼ばれる
                _logger.LogMessage($"SpreadsheetDocument を閉じる/破棄: {displayFilePath}");
            }
            catch (OpenXmlPackageException oxpe) // ファイル破損などで開けない場合
            {
                _logger.LogMessage($"OpenXMLPackageException (ファイル破損の可能性): {displayFilePath}, エラー: {oxpe.Message}. このファイルはスキップされます。");
                // localResults は空のまま返され、このファイルからは結果なしとなる
            }
            catch (OperationCanceledException) // キャンセルは再スロー
            {
                _logger.LogMessage($"ファイル処理がキャンセルされました (SpreadsheetDocument スコープ): {displayFilePath}");
                throw;
            }
            catch (Exception ex) // その他の予期せぬエラー
            {
                _logger.LogMessage($"Excelファイル処理中の予期せぬエラー (SpreadsheetDocument スコープ): {displayFilePath}, {ex.GetType().Name}: {ex.Message}");
                _logger.LogMessage($"エラー詳細: {ex.StackTrace}");
                // このファイルの処理は失敗するが、全体の検索は継続
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
            bool alreadyFoundHitInFile, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (worksheetPart.Worksheet == null)
            {
                _logger.LogMessage($"Worksheet が null です。シート: '{sheetName}', ファイル: {Path.GetFileName(filePath)}");
                return alreadyFoundHitInFile;
            }

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                _logger.LogMessage($"SheetData が null です。シート: '{sheetName}', ファイル: {Path.GetFileName(filePath)}");
                return alreadyFoundHitInFile;
            }

            bool foundInThisSheet = false;

            foreach (Row row in sheetData.Elements<Row>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (firstHitOnly && (alreadyFoundHitInFile || foundInThisSheet)) break;

                foreach (Cell cell in row.Elements<Cell>())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (firstHitOnly && (alreadyFoundHitInFile || foundInThisSheet)) break;

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
                            CellValue = cellValue // TruncateString は表示時に行う方が良いかもしれない
                        };

                        pendingResults.Enqueue(result); // UI表示とメインリストへの追加用
                        localResults.Add(result);       // このファイル内での結果リスト
                        foundInThisSheet = true;

                        _logger.LogMessage($"セル内一致: {Path.GetFileName(filePath)} - {sheetName} - {result.CellPosition} - '{TruncateString(result.CellValue)}'");

                        if (firstHitOnly) break;
                    }
                }
            }
            return alreadyFoundHitInFile || foundInThisSheet;
        }

        /// <summary>
        /// ワークシート内の図形を検索
        /// </summary>
        private bool SearchInShapes(
            string filePath, string sheetName, WorksheetPart worksheetPart,
            string keyword, bool useRegex, Regex regex,
            ConcurrentQueue<SearchResult> pendingResults, List<SearchResult> localResults,
            bool firstHitOnly, bool alreadyFoundHitInFile, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null || drawingsPart.WorksheetDrawing == null)
            {
                 // 図形がない場合はログ出力しない（通常のケースなので冗長になる）
                // _logger.LogMessage($"DrawingsPart が null または WorksheetDrawing が null です。シート: '{sheetName}', ファイル: {Path.GetFileName(filePath)}");
                return alreadyFoundHitInFile;
            }
            _logger.LogMessage($"図形パーツ処理中: シート '{sheetName}', ファイル: {Path.GetFileName(filePath)}");


            bool foundInThisSheet = false;

            foreach (var twoCellAnchor in drawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (firstHitOnly && (alreadyFoundHitInFile || foundInThisSheet)) break;

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
                                CellPosition = "図形内", // より詳細な位置情報が必要なら追加検討
                                CellValue = shapeText    // TruncateString は表示時に
                            };
                            pendingResults.Enqueue(result);
                            localResults.Add(result);
                            foundInThisSheet = true;
                            _logger.LogMessage($"図形内一致 (Shape): {Path.GetFileName(filePath)} - {sheetName} - '{TruncateString(result.CellValue)}'");
                            if (firstHitOnly) break;
                        }
                    }
                }

                if (firstHitOnly && (alreadyFoundHitInFile || foundInThisSheet)) break;

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
                                CellPosition = "図形内 (GF)", // GraphicFrame
                                CellValue = frameText      // TruncateString は表示時に
                            };
                            pendingResults.Enqueue(result);
                            localResults.Add(result);
                            foundInThisSheet = true;
                            _logger.LogMessage($"図形内一致 (GraphicFrame): {Path.GetFileName(filePath)} - {sheetName} - '{TruncateString(result.CellValue)}'");
                            if (firstHitOnly) break;
                        }
                    }
                }
            }
            return alreadyFoundHitInFile || foundInThisSheet;
        }

        /// <summary>
        /// ワークシート名を取得
        /// </summary>
        private string GetSheetName(WorkbookPart workbookPart, WorksheetPart worksheetPart)
        {
            // worksheetPart からシート名を取得するより信頼性の高い方法
            // workbookPart と worksheetPart の関連付け (relationship ID) を使用して、Workbook の中の Sheet 要素から名前を見つける
            string relationshipId = workbookPart.GetIdOfPart(worksheetPart);
            Sheet sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault(s => s.Id != null && s.Id.Value == relationshipId);
            return sheet?.Name?.Value;
        }

        /// <summary>
        /// セルの値を取得
        /// </summary>
        private string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell == null || cell.CellValue == null)
                return string.Empty;

            string value = cell.CellValue.InnerText;

            // DataTypeがSharedStringの場合、共有文字列テーブルから実際の値を取得
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (sharedStringTable != null && int.TryParse(value, out int ssid))
                {
                    // SharedStringItem ssi = sharedStringTable.ChildElements[ssid] as SharedStringItem;
                    // より安全な取得方法
                    SharedStringItem ssi = sharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(ssid);
                    if (ssi != null)
                    {
                        // Text 要素または RichTextRun 要素内の Text 要素の値を連結する
                        // 単純な .InnerText は書式情報を含む場合があるため、Text要素を直接取得する
                        return string.Concat(ssi.Elements<Text>().Select(t => t.Text))
                               + string.Concat(ssi.Elements<Run>()
                                   .SelectMany(r => r.Elements<Text>())
                                   .Select(t => t.Text));
                    }
                }
                return string.Empty; // 共有文字列が見つからない場合
            }
            // DataTypeがBooleanの場合、0か1で返ってくるのでbool値に変換
            else if (cell.DataType != null && cell.DataType.Value == CellValues.Boolean)
            {
                return value == "1" ? "TRUE" : (value == "0" ? "FALSE" : value);
            }
            // DataTypeがDateの場合、OpenXMLのシリアル値を返すので、そのまま文字列として扱うか、DateTimeに変換するか検討
            // ここではそのまま文字列として返す
            // if (cell.DataType != null && cell.DataType.Value == CellValues.Date) { ... }


            return value;
        }

        /// <summary>
        /// セル参照（例：A1）を取得
        /// </summary>
        private string GetCellReference(Cell cell)
        {
            return cell?.CellReference?.Value ?? string.Empty;
        }

        /// <summary>
        /// 文字列を指定の長さに切り詰める (ログ表示用)
        /// </summary>
        private string TruncateString(string value, int maxLength = 100) // ログ用なので短めに
        {
            if (string.IsNullOrEmpty(value)) return value;
            string cleanedValue = value.Replace("\n", " ").Replace("\r", " "); // 改行をスペースに置換
            return cleanedValue.Length <= maxLength ? cleanedValue : cleanedValue.Substring(0, maxLength) + "...";
        }

        /// <summary>
        /// 現在のプロセスのワーキングセットメモリ使用量を取得 (MB単位)
        /// </summary>
        private double GetCurrentMemoryUsageMB()
        {
            // using System.Diagnostics; が必要
            Process currentProcess = Process.GetCurrentProcess();
            currentProcess.Refresh(); // 最新の値を取得するためにRefreshを呼び出す
            return currentProcess.WorkingSet64 / (1024.0 * 1024.0);
        }
    }
}