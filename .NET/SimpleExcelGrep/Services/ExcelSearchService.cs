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

        // プロセスのワーキングセットサイズを設定するためのWin32 API
        [System.Runtime.InteropServices.DllImport( "kernel32.dll" )]
        private static extern bool SetProcessWorkingSetSize ( IntPtr proc, int min, int max );

        /// <summary>
        /// ExcelSearchServiceのコンストラクタ
        /// </summary>
        /// <param name="logger">ログサービス</param>
        public ExcelSearchService ( LogService logger )
        {
            _logger = logger;
            _shapeTextExtractor = new ShapeTextExtractor();
        }

        /// <summary>
        /// 指定されたフォルダ内のExcelファイルを検索（バッチ処理完全実装版）
        /// </summary>
        public async Task<List<SearchResult>> SearchExcelFilesAsync (
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
            CancellationToken cancellationToken )
        {
            _logger.LogMessage( $"SearchExcelFilesAsync 開始: フォルダ={folderPath}" );

            List<SearchResult> finalResults = new List<SearchResult>();

            try
            {
                string [] allExcelFiles = Directory.GetFiles( folderPath, "*.xlsx", SearchOption.AllDirectories )
                                         .Concat( Directory.GetFiles( folderPath, "*.xlsm", SearchOption.AllDirectories ) )
                                         .ToArray();

                _logger.LogMessage( $"{allExcelFiles.Length} 個のExcelファイル(.xlsx, .xlsm)が見つかりました" );

                int totalFiles = allExcelFiles.Length;
                int processedFiles = 0;

                // ファイルを適切なサイズのバッチに分割
                // 大量ファイル処理のため、バッチサイズを小さく設定
                int batchSize = CalculateSafeBatchSize( maxParallelism );
                int totalBatches = ( totalFiles + batchSize - 1 ) / batchSize;

                _logger.LogMessage( $"処理を {totalBatches} バッチに分割します (1バッチ {batchSize} ファイル)" );

                // 各バッチを順次処理
                for ( int batchIndex = 0; batchIndex < totalBatches; batchIndex++ )
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    int startIndex = batchIndex * batchSize;
                    int count = Math.Min( batchSize, totalFiles - startIndex );
                    string [] batchFiles = new string [ count ];
                    Array.Copy( allExcelFiles, startIndex, batchFiles, 0, count );

                    statusUpdateCallback( $"バッチ {batchIndex + 1}/{totalBatches} 処理中... ({processedFiles}/{totalFiles} 完了)" );
                    _logger.LogMessage( $"バッチ {batchIndex + 1}/{totalBatches} 処理開始 (ファイル {startIndex + 1}-{startIndex + count}/{totalFiles})" );

                    // メモリリーク検出用
                    long memoryBefore = Process.GetCurrentProcess().WorkingSet64;

                    // 一時結果を格納するリスト
                    List<SearchResult> batchResults = new List<SearchResult>();
                    ConcurrentBag<SearchResult> concurrentBatchResults = new ConcurrentBag<SearchResult>();

                    var parallelOptions = new ParallelOptions
                    {
                        MaxDegreeOfParallelism = maxParallelism,
                        CancellationToken = cancellationToken
                    };

                    // バッチ内の並列処理
                    await Task.Run( () =>
                    {
                        Parallel.ForEach( batchFiles, parallelOptions, ( filePath ) =>
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            // 無視キーワードチェック
                            if ( ignoreKeywords.Any( k => filePath.IndexOf( k, StringComparison.OrdinalIgnoreCase ) >= 0 ) )
                            {
                                Interlocked.Increment( ref processedFiles );
                                return;
                            }

                            // ファイルサイズチェック
                            if ( ignoreFileSizeMB > 0 )
                            {
                                try
                                {
                                    FileInfo fileInfo = new FileInfo( filePath );
                                    double fileSizeMB = ( double ) fileInfo.Length / ( 1024 * 1024 );
                                    if ( fileSizeMB > ignoreFileSizeMB )
                                    {
                                        Interlocked.Increment( ref processedFiles );
                                        return;
                                    }
                                }
                                catch ( Exception )
                                {
                                    // エラーは無視
                                }
                            }

                            try
                            {
                                string extension = Path.GetExtension( filePath ).ToLowerInvariant();

                                if ( extension == ".xlsx" || extension == ".xlsm" )
                                {
                                    List<SearchResult> fileResults = SearchInXlsxFile(
                                        filePath, keyword, useRegex, regex, resultQueue,
                                        firstHitOnly, searchShapes, cancellationToken );

                                    // 結果を追加
                                    foreach ( var result in fileResults )
                                    {
                                        concurrentBatchResults.Add( result );
                                        if ( isRealTimeDisplay )
                                        {
                                            resultQueue.Enqueue( result );
                                        }
                                    }
                                }
                            }
                            catch ( OperationCanceledException )
                            {
                                throw;
                            }
                            catch ( Exception ex )
                            {
                                _logger.LogMessage( $"ファイル処理エラー: {filePath}, {ex.Message}" );
                            }
                            finally
                            {
                                int currentProcessed = Interlocked.Increment( ref processedFiles );
                                if ( currentProcessed % 20 == 0 || currentProcessed == totalFiles )
                                {
                                    statusUpdateCallback( $"バッチ {batchIndex + 1}/{totalBatches} 処理中... ({currentProcessed}/{totalFiles} 完了)" );
                                }
                            }
                        } );

                        // ConcurrentBagから通常のリストに移す
                        batchResults.AddRange( concurrentBatchResults );
                    }, cancellationToken );

                    // バッチ結果を最終結果に追加
                    finalResults.AddRange( batchResults );

                    // 結果を表示ためにキューに追加（リアルタイム表示でない場合）
                    if ( !isRealTimeDisplay )
                    {
                        foreach ( var result in batchResults )
                        {
                            resultQueue.Enqueue( result );
                        }
                        // 一度にすべての結果を表示するとUIがフリーズするのを防ぐ
                        await Task.Delay( 100, cancellationToken );
                    }

                    // バッチ処理完了後の強制的なメモリクリーンアップ
                    _logger.LogMessage( $"バッチ {batchIndex + 1} 処理完了。メモリクリーンアップを実行..." );
                    batchResults.Clear();
                    concurrentBatchResults = null;

                    // 強制的なメモリクリーンアップ
                    ForceMemoryCleanup();

                    // メモリリーク検出
                    long memoryAfter = Process.GetCurrentProcess().WorkingSet64;
                    double memoryDiffMB = ( memoryAfter - memoryBefore ) / ( 1024.0 * 1024.0 );

                    _logger.LogMessage( $"バッチ処理後のメモリ差分: {memoryDiffMB:F2}MB (処理前: {memoryBefore / ( 1024 * 1024 )}MB, 処理後: {memoryAfter / ( 1024 * 1024 )}MB)" );

                    // メモリ増加が大きい場合は警告
                    if ( memoryDiffMB > 500 ) // 500MB以上増加している場合
                    {
                        _logger.LogMessage( $"警告: メモリリークの可能性があります。バッチ処理後にメモリが {memoryDiffMB:F2}MB 増加しています" );

                        // さらに強力なクリーンアップを試みる
                        ForceMemoryCleanup();
                        ForceMemoryCleanup();
                    }
                }
            }
            catch ( OperationCanceledException )
            {
                _logger.LogMessage( "SearchExcelFilesAsync内のタスクがキャンセルされました。" );
                throw;
            }
            catch ( Exception ex )
            {
                _logger.LogMessage( $"SearchExcelFilesAsync内のタスクで予期せぬエラー: {ex.Message}" );
                throw;
            }
            finally
            {
                // 終了時には完全なメモリクリーンアップを実行
                _logger.LogMessage( "SearchExcelFilesAsync 終了時のメモリクリーンアップ実行" );
                ForceMemoryCleanup();
            }

            _logger.LogMessage( $"SearchExcelFilesAsync 完了: {finalResults.Count}件の結果" );
            return finalResults;
        }

        /// <summary>
        /// 単一のXLSX/XLSMファイルを検索（完全クリーンアップ版）
        /// </summary>
        private List<SearchResult> SearchInXlsxFile (
            string filePath,
            string keyword,
            bool useRegex,
            Regex regex,
            ConcurrentQueue<SearchResult> pendingResults,
            bool firstHitOnly,
            bool searchShapes,
            CancellationToken cancellationToken )
        {
            List<SearchResult> localResults = new List<SearchResult>();
            bool foundHitInFile = false;
            SpreadsheetDocument spreadsheetDocument = null;

            try
            {
                // 各ファイル処理前にGCを実行
                if ( Process.GetCurrentProcess().WorkingSet64 > 4L * 1024 * 1024 * 1024 ) // 4GB超
                {
                    GC.Collect( 0, GCCollectionMode.Optimized, false );
                }

                spreadsheetDocument = SpreadsheetDocument.Open( filePath, false );

                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                if ( workbookPart == null )
                {
                    spreadsheetDocument.Dispose();
                    return localResults;
                }

                SharedStringTablePart sharedStringTablePart = null;
                SharedStringTable sharedStringTable = null;

                try
                {
                    sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    sharedStringTable = sharedStringTablePart?.SharedStringTable;

                    // 各ワークシートを処理
                    foreach ( WorksheetPart worksheetPart in workbookPart.WorksheetParts )
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if ( firstHitOnly && foundHitInFile ) break;

                        string sheetName = GetSheetName( workbookPart, worksheetPart ) ?? "不明なシート";

                        // 1. セル内のテキスト検索
                        foundHitInFile = SearchInCells(
                            filePath, sheetName, worksheetPart, sharedStringTable,
                            keyword, useRegex, regex, pendingResults,
                            localResults, firstHitOnly, foundHitInFile, cancellationToken );

                        // 2. 図形内のテキスト検索
                        if ( searchShapes )
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if ( firstHitOnly && foundHitInFile ) break;

                            foundHitInFile = SearchInShapes(
                                filePath, sheetName, worksheetPart,
                                keyword, useRegex, regex, pendingResults,
                                localResults, firstHitOnly, foundHitInFile, cancellationToken );
                        }

                        // 各WorksheetPartの使用後に完全クリーンアップ
                        CompleteCleanupWorksheetPart( worksheetPart );
                    }

                    // 共有文字列テーブルの完全クリーンアップ
                    if ( sharedStringTable != null )
                    {
                        CompleteCleanupSharedStringTable( sharedStringTable );
                    }

                    // WorkbookPartのクリーンアップ
                    CompleteCleanupWorkbookPart( workbookPart );
                }
                finally
                {
                    // 参照をnullに設定して解放を助ける
                    sharedStringTable = null;
                    sharedStringTablePart = null;
                    workbookPart = null;
                }
            }
            catch ( OperationCanceledException )
            {
                throw;
            }
            catch ( Exception ex )
            {
                _logger.LogMessage( $"Excel処理エラー: {filePath}, {ex.GetType().Name}: {ex.Message}" );
            }
            finally
            {
                // SpreadsheetDocumentの完全な解放
                if ( spreadsheetDocument != null )
                {
                    try
                    {
                        spreadsheetDocument.Dispose();
                    }
                    catch
                    {
                        // エラー無視
                    }
                    spreadsheetDocument = null;
                }
            }

            return localResults;
        }

        /// <summary>
        /// ワークシート内のセルを検索
        /// </summary>
        private bool SearchInCells (
            string filePath, string sheetName, WorksheetPart worksheetPart,
            SharedStringTable sharedStringTable, string keyword, bool useRegex,
            Regex regex, ConcurrentQueue<SearchResult> pendingResults,
            List<SearchResult> localResults, bool firstHitOnly,
            bool foundHitInFile, CancellationToken cancellationToken )
        {
            if ( worksheetPart.Worksheet == null ) return foundHitInFile;

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if ( sheetData == null ) return foundHitInFile;

            foreach ( Row row in sheetData.Elements<Row>() )
            {
                cancellationToken.ThrowIfCancellationRequested();
                if ( firstHitOnly && foundHitInFile ) break;

                foreach ( Cell cell in row.Elements<Cell>() )
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if ( firstHitOnly && foundHitInFile ) break;

                    string cellValue = GetCellValue( cell, sharedStringTable );
                    if ( string.IsNullOrEmpty( cellValue ) ) continue;

                    bool isMatch = useRegex && regex != null ?
                                   regex.IsMatch( cellValue ) :
                                   cellValue.IndexOf( keyword, StringComparison.OrdinalIgnoreCase ) >= 0;

                    if ( isMatch )
                    {
                        SearchResult result = new SearchResult
                        {
                            FilePath = filePath,
                            SheetName = sheetName,
                            CellPosition = GetCellReference( cell ),
                            CellValue = cellValue
                        };

                        pendingResults.Enqueue( result );
                        localResults.Add( result );
                        foundHitInFile = true;

                        _logger.LogMessage( $"セル内一致: {filePath} - {sheetName} - {result.CellPosition} - '{TruncateString( result.CellValue )}'" );

                        if ( firstHitOnly ) break;
                    }
                }
            }

            return foundHitInFile;
        }

        /// <summary>
        /// ワークシート内の図形を検索
        /// </summary>
        private bool SearchInShapes (
            string filePath, string sheetName, WorksheetPart worksheetPart,
            string keyword, bool useRegex, Regex regex,
            ConcurrentQueue<SearchResult> pendingResults, List<SearchResult> localResults,
            bool firstHitOnly, bool foundHitInFile, CancellationToken cancellationToken )
        {
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart;
            if ( drawingsPart == null || drawingsPart.WorksheetDrawing == null ) return foundHitInFile;

            foreach ( var twoCellAnchor in drawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>() )
            {
                cancellationToken.ThrowIfCancellationRequested();
                if ( firstHitOnly && foundHitInFile ) break;

                // Shapeからテキスト検索
                var shape = twoCellAnchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().FirstOrDefault();
                if ( shape != null && shape.TextBody != null )
                {
                    string shapeText = _shapeTextExtractor.GetTextFromShapeTextBody( shape.TextBody );
                    if ( !string.IsNullOrEmpty( shapeText ) )
                    {
                        bool isMatch = useRegex && regex != null ?
                                      regex.IsMatch( shapeText ) :
                                      shapeText.IndexOf( keyword, StringComparison.OrdinalIgnoreCase ) >= 0;

                        if ( isMatch )
                        {
                            SearchResult result = new SearchResult
                            {
                                FilePath = filePath,
                                SheetName = sheetName,
                                CellPosition = "図形内",
                                CellValue = TruncateString( shapeText )
                            };

                            pendingResults.Enqueue( result );
                            localResults.Add( result );
                            foundHitInFile = true;

                            _logger.LogMessage( $"図形内一致 (Shape): {filePath} - {sheetName} - '{result.CellValue}'" );

                            if ( firstHitOnly ) break;
                        }
                    }
                }

                // GraphicFrameからテキスト検索
                var graphicFrame = twoCellAnchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>().FirstOrDefault();
                if ( graphicFrame != null )
                {
                    string frameText = _shapeTextExtractor.GetTextFromGraphicFrame( graphicFrame );
                    if ( !string.IsNullOrEmpty( frameText ) )
                    {
                        bool isMatch = useRegex && regex != null ?
                                      regex.IsMatch( frameText ) :
                                      frameText.IndexOf( keyword, StringComparison.OrdinalIgnoreCase ) >= 0;

                        if ( isMatch )
                        {
                            SearchResult result = new SearchResult
                            {
                                FilePath = filePath,
                                SheetName = sheetName,
                                CellPosition = "図形内 (GF)",
                                CellValue = TruncateString( frameText )
                            };

                            pendingResults.Enqueue( result );
                            localResults.Add( result );
                            foundHitInFile = true;

                            _logger.LogMessage( $"図形内一致 (GraphicFrame): {filePath} - {sheetName} - '{result.CellValue}'" );

                            if ( firstHitOnly ) break;
                        }
                    }
                }
            }

            return foundHitInFile;
        }

        /// <summary>
        /// ワークシート名を取得
        /// </summary>
        private string GetSheetName ( WorkbookPart workbookPart, WorksheetPart worksheetPart )
        {
            string sheetId = workbookPart.GetIdOfPart( worksheetPart );
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault( s => s.Id?.Value == sheetId );
            return sheet?.Name?.Value;
        }

        /// <summary>
        /// セルの値を取得
        /// </summary>
        private string GetCellValue ( Cell cell, SharedStringTable sharedStringTable )
        {
            if ( cell == null || cell.CellValue == null )
                return string.Empty;

            string cellValueStr = cell.CellValue.InnerText;

            if ( cell.DataType != null && cell.DataType.Value == CellValues.SharedString && sharedStringTable != null )
            {
                if ( int.TryParse( cellValueStr, out int ssid ) && ssid >= 0 && ssid < sharedStringTable.ChildElements.Count )
                {
                    SharedStringItem ssi = sharedStringTable.ChildElements [ ssid ] as SharedStringItem;
                    if ( ssi != null )
                    {
                        // Text 要素の値を連結する
                        return string.Concat( ssi.Elements<Text>().Select( t => t.Text ) );
                    }
                }
                return string.Empty; // 共有文字列が見つからない場合
            }
            return cellValueStr;
        }

        /// <summary>
        /// セル参照（例：A1）を取得
        /// </summary>
        private string GetCellReference ( Cell cell )
        {
            return cell.CellReference?.Value ?? string.Empty;
        }

        /// <summary>
        /// 文字列を指定の長さに切り詰める
        /// </summary>
        private string TruncateString ( string value, int maxLength = 255 )
        {
            if ( string.IsNullOrEmpty( value ) ) return value;
            return value.Length <= maxLength ? value : value.Substring( 0, maxLength ) + "...";
        }

        /// <summary>
        /// WorksheetPartのリソースを解放
        /// </summary>
        private void CleanupWorksheetPart ( WorksheetPart worksheetPart )
        {
            if ( worksheetPart == null ) return;

            try
            {
                // シートデータのクリーンアップ
                if ( worksheetPart.Worksheet != null )
                {
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    if ( sheetData != null && sheetData.HasChildren )
                    {
                        foreach ( var row in sheetData.Elements<Row>() )
                        {
                            if ( row.HasChildren )
                            {
                                foreach ( var cell in row.Elements<Cell>() )
                                {
                                    if ( cell.CellValue != null )
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
                    foreach ( var element in worksheetPart.Worksheet.ChildElements )
                    {
                        if ( element.HasChildren )
                        {
                            element.RemoveAllChildren();
                        }
                    }
                }

                // 図形データのクリーンアップ
                if ( worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null )
                {
                    var drawing = worksheetPart.DrawingsPart.WorksheetDrawing;

                    // TwoCellAnchorの子要素をクリーンアップ
                    foreach ( var anchor in drawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>() )
                    {
                        if ( anchor.HasChildren )
                        {
                            var shapes = anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().ToList();
                            foreach ( var shape in shapes )
                            {
                                if ( shape.TextBody != null && shape.TextBody.HasChildren )
                                {
                                    shape.TextBody.RemoveAllChildren();
                                }
                                shape.RemoveAllChildren();
                            }

                            var graphicFrames = anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>().ToList();
                            foreach ( var frame in graphicFrames )
                            {
                                if ( frame.Graphic != null && frame.Graphic.GraphicData != null )
                                {
                                    frame.Graphic.GraphicData.RemoveAllChildren();
                                }
                                if ( frame.Graphic != null )
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
            catch ( Exception ex )
            {
                _logger.LogMessage( $"WorksheetPartクリーンアップエラー: {ex.Message}" );
            }
        }

        /// <summary>
        /// WorkbookPartのリソースを解放
        /// </summary>
        private void CleanupWorkbookPart ( WorkbookPart workbookPart )
        {
            if ( workbookPart == null ) return;

            try
            {
                // Sheets, DefinedNames, BookViewsのクリーンアップ
                if ( workbookPart.Workbook != null )
                {
                    var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                    if ( sheets != null && sheets.HasChildren )
                    {
                        sheets.RemoveAllChildren();
                    }

                    var definedNames = workbookPart.Workbook.GetFirstChild<DefinedNames>();
                    if ( definedNames != null && definedNames.HasChildren )
                    {
                        definedNames.RemoveAllChildren();
                    }

                    var bookViews = workbookPart.Workbook.GetFirstChild<BookViews>();
                    if ( bookViews != null && bookViews.HasChildren )
                    {
                        bookViews.RemoveAllChildren();
                    }
                }

                // Styles, Theme, Calculation Propertiesなどの追加部品もクリーンアップ
                if ( workbookPart.WorkbookStylesPart != null )
                {
                    var styles = workbookPart.WorkbookStylesPart.Stylesheet;
                    if ( styles != null )
                    {
                        // スタイル情報をクリア
                        styles.RemoveAllChildren();
                    }
                }
            }
            catch ( Exception ex )
            {
                _logger.LogMessage( $"WorkbookPartクリーンアップエラー: {ex.Message}" );
            }
        }

        /// <summary>
        /// 最適なバッチサイズを計算
        /// </summary>
        private int CalculateOptimalBatchSize ( string [] files, double ignoreFileSizeMB, int maxParallelism )
        {
            // ファイル数が少ない場合はバッチ分けしない
            if ( files.Length <= maxParallelism * 2 )
            {
                return files.Length;
            }

            try
            {
                // サンプルファイルのサイズを取得してバッチサイズを推定
                long totalSampleSize = 0;
                int sampleCount = Math.Min( 10, files.Length );

                for ( int i = 0; i < sampleCount; i++ )
                {
                    try
                    {
                        FileInfo fileInfo = new FileInfo( files [ i ] );
                        totalSampleSize += fileInfo.Length;
                    }
                    catch
                    {
                        // ファイルアクセスエラーは無視
                    }
                }

                // 平均ファイルサイズ (バイト)
                double avgFileSizeMB = ( double ) totalSampleSize / ( sampleCount * 1024 * 1024 );

                // 使用可能メモリを考慮したバッチサイズ (目安: 1ファイルあたり平均サイズの5倍のメモリを使用すると仮定)
                long availableMemory = GC.GetTotalMemory( false );
                double memoryPerFileMB = Math.Max( 1, avgFileSizeMB * 5 ); // 最低1MBと仮定

                // 利用可能メモリの80%を使用すると仮定
                double usableMemoryMB = Math.Max( 100, availableMemory / ( 1024 * 1024 ) * 0.8 );

                // バッチサイズを計算 (メモリベース)
                int memoryBasedBatchSize = ( int ) Math.Max( 1, Math.Min( 100, usableMemoryMB / memoryPerFileMB ) );

                // 並列処理数ベースのバッチサイズ (通常は並列数の2-5倍がバランスが良い)
                int threadBasedBatchSize = maxParallelism * 3;

                // 最終的なバッチサイズを決定（小さい方を選択）
                int batchSize = Math.Min( memoryBasedBatchSize, threadBasedBatchSize );

                // 極端に小さいか大きい値にならないよう制限
                return Math.Max( 5, Math.Min( 50, batchSize ) );
            }
            catch ( Exception ex )
            {
                _logger.LogMessage( $"バッチサイズ計算エラー: {ex.Message}" );
                return maxParallelism * 2; // デフォルト値
            }
        }

        /// <summary>
        /// WorksheetPartの軽量クリーンアップ
        /// </summary>
        private void LightCleanupWorksheetPart ( WorksheetPart worksheetPart )
        {
            // 重要なリソースの解放のみに留める
            // ほとんどの場合は明示的な解放は不要（ガベージコレクション任せが速い）

            // 大量のメモリを消費している可能性がある図形データのみクリーンアップ
            try
            {
                if ( worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null )
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
        private void BalancedCleanupWorksheetPart ( WorksheetPart worksheetPart )
        {
            if ( worksheetPart == null ) return;

            try
            {
                // 最も重要な図形データ部分のクリーンアップ
                if ( worksheetPart.DrawingsPart != null && worksheetPart.DrawingsPart.WorksheetDrawing != null )
                {
                    var drawing = worksheetPart.DrawingsPart.WorksheetDrawing;

                    // TwoCellAnchorの子要素のうち、大きなデータを持つものだけクリーンアップ
                    foreach ( var anchor in drawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>() )
                    {
                        if ( anchor.HasChildren )
                        {
                            // 特にテキストを持つShapeやGraphicFrameは確実にクリーンアップ
                            var shapes = anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().ToList();
                            foreach ( var shape in shapes )
                            {
                                if ( shape.TextBody != null && shape.TextBody.HasChildren )
                                {
                                    shape.TextBody.RemoveAllChildren();
                                }
                            }

                            var graphicFrames = anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>().ToList();
                            foreach ( var frame in graphicFrames )
                            {
                                if ( frame.Graphic != null && frame.Graphic.GraphicData != null )
                                {
                                    frame.Graphic.GraphicData.RemoveAllChildren();
                                }
                            }
                        }
                    }
                }

                // シートデータは条件付きでクリーンアップ
                if ( worksheetPart.Worksheet != null )
                {
                    // セルの値だけをクリア（構造は維持）
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    if ( sheetData != null && sheetData.HasChildren )
                    {
                        foreach ( var row in sheetData.Elements<Row>() )
                        {
                            foreach ( var cell in row.Elements<Cell>() )
                            {
                                if ( cell.CellValue != null )
                                {
                                    // セル値のテキスト内容だけクリア
                                    cell.CellValue.Text = string.Empty;
                                }
                            }
                        }
                    }
                }
            }
            catch ( Exception ex )
            {
                // エラーをログに記録するが、処理は継続
                _logger.LogMessage( $"WorksheetPartクリーンアップエラー: {ex.Message}" );
            }
        }

        /// <summary>
        /// WorkbookPartのバランスのとれたクリーンアップ
        /// </summary>
        private void BalancedCleanupWorkbookPart ( WorkbookPart workbookPart )
        {
            // 最低限のクリーンアップのみ実行
            if ( workbookPart?.WorkbookStylesPart?.Stylesheet != null )
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
        private void SynchronizeResults ( List<List<SearchResult>> threadLocalResults, List<SearchResult> mainResults, ConcurrentQueue<SearchResult> resultQueue )
        {
            lock ( mainResults )
            {
                for ( int i = 0; i < threadLocalResults.Count; i++ )
                {
                    lock ( threadLocalResults [ i ] )
                    {
                        // 新しい結果のみ追加
                        int startCount = mainResults.Count;
                        mainResults.AddRange( threadLocalResults [ i ] );

                        // すでに同期した結果はクリア
                        threadLocalResults [ i ].Clear();

                        _logger.LogMessage( $"結果同期: スレッド {i} から {mainResults.Count - startCount} 件追加" );
                    }
                }
            }
        }

        /// <summary>
        /// WorksheetPartの完全クリーンアップ
        /// </summary>
        private void CompleteCleanupWorksheetPart ( WorksheetPart worksheetPart )
        {
            if ( worksheetPart == null ) return;

            try
            {
                // 図形データの完全クリーンアップ
                if ( worksheetPart.DrawingsPart != null )
                {
                    if ( worksheetPart.DrawingsPart.WorksheetDrawing != null )
                    {
                        // 全ての子要素を切り離し
                        worksheetPart.DrawingsPart.WorksheetDrawing.RemoveAllChildren();
                        worksheetPart.DrawingsPart.WorksheetDrawing = null;
                    }

                    // 全てのImagePartを切り離し
                    var imageParts = worksheetPart.DrawingsPart.ImageParts.ToList();
                    foreach ( var part in imageParts )
                    {
                        worksheetPart.DrawingsPart.DeletePart( part );
                    }

                    // DrawingsPartの切り離し
                    worksheetPart.DeletePart( worksheetPart.DrawingsPart );
                }

                // シートデータの完全クリーンアップ
                if ( worksheetPart.Worksheet != null )
                {
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    if ( sheetData != null )
                    {
                        sheetData.RemoveAllChildren();
                    }

                    // その他のシート要素も削除
                    worksheetPart.Worksheet.RemoveAllChildren();

                    // 参照を解除
                    worksheetPart.Worksheet = null;
                }

                // その他のパーツも全て削除
                var parts = worksheetPart.Parts.Select( p => p.OpenXmlPart ).ToList();
                foreach ( var part in parts )
                {
                    if ( part != worksheetPart.DrawingsPart ) // 既に処理済みのものを避ける
                    {
                        try
                        {
                            worksheetPart.DeletePart( part );
                        }
                        catch
                        {
                            // エラー無視
                        }
                    }
                }
            }
            catch ( Exception ex )
            {
                _logger.LogMessage( $"WorksheetPart完全クリーンアップエラー: {ex.Message}" );
            }
        }

        /// <summary>
        /// 共有文字列テーブルの完全クリーンアップ
        /// </summary>
        private void CompleteCleanupSharedStringTable ( SharedStringTable sharedStringTable )
        {
            if ( sharedStringTable == null ) return;

            try
            {
                // 全ての子要素を個別に削除
                var items = sharedStringTable.Elements<SharedStringItem>().ToList();
                foreach ( var item in items )
                {
                    item.RemoveAllChildren();
                    sharedStringTable.RemoveChild( item );
                }

                // テーブル自体の全ての子要素を削除
                sharedStringTable.RemoveAllChildren();
            }
            catch ( Exception ex )
            {
                _logger.LogMessage( $"SharedStringTableクリーンアップエラー: {ex.Message}" );
            }
        }

        /// <summary>
        /// WorkbookPartの完全クリーンアップ
        /// </summary>
        private void CompleteCleanupWorkbookPart ( WorkbookPart workbookPart )
        {
            if ( workbookPart == null ) return;

            try
            {
                // ワークブック内の各種要素をクリーンアップ
                if ( workbookPart.Workbook != null )
                {
                    var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                    if ( sheets != null )
                    {
                        sheets.RemoveAllChildren();
                    }

                    // WorkbookPropertiesなどもクリーンアップ
                    workbookPart.Workbook.RemoveAllChildren();
                }

                // スタイル部分のクリーンアップ
                if ( workbookPart.WorkbookStylesPart != null )
                {
                    if ( workbookPart.WorkbookStylesPart.Stylesheet != null )
                    {
                        workbookPart.WorkbookStylesPart.Stylesheet.RemoveAllChildren();
                    }
                    workbookPart.DeletePart( workbookPart.WorkbookStylesPart );
                }

                // テーマなどその他パーツのクリーンアップ
                var parts = workbookPart.Parts.Select( p => p.OpenXmlPart ).ToList();
                foreach ( var part in parts )
                {
                    // WorksheetPartsはすでに別途処理済みなので除外
                    if ( !( part is WorksheetPart ) && !( part is SharedStringTablePart ) )
                    {
                        try
                        {
                            workbookPart.DeletePart( part );
                        }
                        catch
                        {
                            // エラー無視
                        }
                    }
                }

                // 最後に参照を解除
                workbookPart.Workbook = null;
            }
            catch ( Exception ex )
            {
                _logger.LogMessage( $"WorkbookPartクリーンアップエラー: {ex.Message}" );
            }
        }

        /// <summary>
        /// 安全なバッチサイズを計算
        /// </summary>
        private int CalculateSafeBatchSize ( int maxParallelism )
        {
            // システムメモリに基づいてバッチサイズを調整
            try
            {
                // 利用可能物理メモリを取得（MB）
                long availableMemoryMB = 0;

                try
                {
                    using ( var pc = new PerformanceCounter( "Memory", "Available MBytes" ) )
                    {
                        availableMemoryMB = (long)pc.NextValue();
                    }
                }
                catch
                {
                    // PerformanceCounterが使えない場合は代替手段
                    availableMemoryMB = GC.GetTotalMemory(false) / (1024 * 1024);
                }

                // 利用可能メモリに基づくバッチサイズ（1ファイルあたり約10MBと仮定）
                int memoryBasedSize = ( int ) ( availableMemoryMB / 20 ); // 利用可能メモリの1/20

                // 並列数に基づくバッチサイズ（1スレッドあたり5-10ファイル程度が最適）
                int threadBasedSize = maxParallelism * 8;

                // メモリベースとスレッドベースの小さい方を選択
                int batchSize = Math.Min( memoryBasedSize, threadBasedSize );

                // 極端に小さな値や大きな値を避ける
                batchSize = Math.Max( maxParallelism * 2, Math.Min( 200, batchSize ) );

                _logger.LogMessage( $"計算されたバッチサイズ: {batchSize} (利用可能メモリ: {availableMemoryMB}MB)" );

                return batchSize;
            }
            catch ( Exception ex )
            {
                _logger.LogMessage( $"バッチサイズ計算エラー: {ex.Message}" );
                return maxParallelism * 4; // デフォルト値
            }
        }

        /// <summary>
        /// 強制的なメモリクリーンアップを実行
        /// </summary>
        private void ForceMemoryCleanup ()
        {
            try
            {
                _logger.LogMessage( $"強制メモリクリーンアップ開始 (現在: {Process.GetCurrentProcess().WorkingSet64 / ( 1024 * 1024 )}MB)" );

                // 段階的なGCを実行
                GC.Collect( 0, GCCollectionMode.Forced, true, true );
                GC.WaitForPendingFinalizers();
                GC.Collect( 1, GCCollectionMode.Forced, true, true );
                GC.WaitForPendingFinalizers();
                GC.Collect( 2, GCCollectionMode.Forced, true, true );
                GC.WaitForPendingFinalizers();

                // プロセスのワーキングセットを明示的に縮小
                try
                {
                    SetProcessWorkingSetSize( Process.GetCurrentProcess().Handle, -1, -1 );
                }
                catch
                {
                    // エラーは無視
                }

                _logger.LogMessage( $"強制メモリクリーンアップ完了 (現在: {Process.GetCurrentProcess().WorkingSet64 / ( 1024 * 1024 )}MB)" );
            }
            catch ( Exception ex )
            {
                _logger.LogMessage( $"メモリクリーンアップエラー: {ex.Message}" );
            }
        }
    }
}