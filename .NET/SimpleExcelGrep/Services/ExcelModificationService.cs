using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
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
            if (!options.RemoveAllFormulas && !options.RemoveAllShapes) return new OtherOperationResult();

            var result = new OtherOperationResult();
            var searchOption = options.IncludeSubDirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var excelFiles = Directory.GetFiles(options.FolderPath, "*.xlsx", searchOption)
                                      .Concat(Directory.GetFiles(options.FolderPath, "*.xlsm", searchOption))
                                      .ToArray();

            result.TotalFiles = excelFiles.Length;

            var logEnabled = options.EnableLogOutput;
            _logger.LogMessage($"その他処理開始: フォルダ={options.FolderPath}, サブフォルダ={options.IncludeSubDirectories}, 無視サイズ(MB)={options.IgnoreFileSizeMb}, 非表示シート対象={options.IncludeInvisibleSheets}, 並列数={options.MaxParallelism}, 図削除={options.RemoveAllShapes}, 数式削除={options.RemoveAllFormulas}", force: logEnabled);

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
                                if (options.RemoveAllShapes)
                                {
                                    fileModified |= RemoveAllShapes(workbookPart, options.IncludeInvisibleSheets, cancellationToken);
                                }
                                if (options.RemoveAllFormulas)
                                {
                                    var sharedStrings = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable;
                                    fileModified |= ReplaceFormulasWithValues(workbookPart, sharedStrings, options.IncludeInvisibleSheets, cancellationToken);
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
    }

    public class OtherOperationResult
    {
        public int TotalFiles { get; set; }
        public int ProcessedFiles { get; set; }
        public int ModifiedFiles { get; set; }
        public int ErrorCount { get; set; }
    }
}
