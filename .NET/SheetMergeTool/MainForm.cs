using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SheetMergeTool
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            lblStatus.Text = "準備完了"; // Initial status
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Excelファイルが含まれるフォルダを選択してください";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtDirPath.Text = dialog.SelectedPath;
                }
            }
        }

        private async void btnProcess_Click(object sender, EventArgs e)
        {
            string folderPath = txtDirPath.Text;
            string startCellRef = txtStartCell.Text.Trim().ToUpper();
            string endCellRef = txtEndCell.Text.Trim().ToUpper();
            bool includeSubdirectories = chkEnableSubDir.Checked; // Read checkbox state

            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                MessageBox.Show("有効なExcelフォルダパスを指定してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!IsValidCellReference(startCellRef))
            {
                MessageBox.Show("有効な開始セル位置を指定してください (例: A1)。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!IsValidCellReference(endCellRef))
            {
                MessageBox.Show("有効な終了セル位置を指定してください (例: D1)。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SetControlsEnabled(false);
            lblStatus.Text = "処理中...";
            txtResults.Clear();
            Application.DoEvents(); 

            var progressLog = new Progress<string>(message => {
                if (txtResults.IsDisposed) return;
                txtResults.AppendText(message + Environment.NewLine);
            });
            var progressStatus = new Progress<string>(status => {
                if (lblStatus.IsDisposed) return;
                lblStatus.Text = status;
            });

            try
            {
                await Task.Run(() => 
                    MergeExcelFilesLogic(folderPath, startCellRef, endCellRef, includeSubdirectories, progressLog, progressStatus) // Pass includeSubdirectories
                );
                ((IProgress<string>)progressStatus).Report("処理完了！");
            }
            catch (Exception ex)
            {
                ((IProgress<string>)progressLog).Report($"エラーが発生しました: {ex.Message}");
                ((IProgress<string>)progressLog).Report($"スタックトレース: {ex.StackTrace}");
                ((IProgress<string>)progressStatus).Report("エラー発生");
                MessageBox.Show($"処理中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetControlsEnabled(true);
            }
        }

        private void SetControlsEnabled(bool enabled)
        {
            txtDirPath.Enabled = enabled;
            txtStartCell.Enabled = enabled;
            txtEndCell.Enabled = enabled;
            chkEnableSubDir.Enabled = enabled; // Manage checkbox state
            btnBrowse.Enabled = enabled;
            btnProcess.Enabled = enabled;
        }

        private void MergeExcelFilesLogic(string folderPath, string startCellUserRef, string endCellUserRef, bool includeSubdirectories,
                                          IProgress<string> logger, IProgress<string> statusUpdater)
        {
            logger.Report("処理を開始します...");
            logger.Report($"対象フォルダ: {folderPath}");
            logger.Report($"サブフォルダを検索: {(includeSubdirectories ? "はい" : "いいえ")}");
            logger.Report($"開始セル: {startCellUserRef}, 終了セル列範囲: {endCellUserRef}");

            string outputFileName = DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + ".xlsx";
            string outputFilePath = Path.Combine(folderPath, outputFileName);

            GetCellRowColumn(startCellUserRef, out uint startRowIndex, out string startColumnName);
            GetCellRowColumn(endCellUserRef, out _, out string endColumnName); 

            int startColumnNum = GetColumnIndexFromName(startColumnName); 
            int endColumnNum = GetColumnIndexFromName(endColumnName);

            if (startColumnNum > endColumnNum)
            {
                logger.Report("警告: 開始セル列が終了セル列より後です。データ列は抽出されません。");
            }
            
            statusUpdater.Report("出力ファイルを準備中...");

            using (SpreadsheetDocument outputDoc = SpreadsheetDocument.Create(outputFilePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = outputDoc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = outputDoc.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet() { Id = outputDoc.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Output" };
                sheets.Append(sheet);

                uint outputCurrentRow = 1;

                Row headerRow = new Row() { RowIndex = outputCurrentRow };
                headerRow.Append(CreateTextCell(GetColumnNameFromIndex(0), outputCurrentRow, "ファイルパス"));
                headerRow.Append(CreateTextCell(GetColumnNameFromIndex(1), outputCurrentRow, "ファイル名"));
                headerRow.Append(CreateTextCell(GetColumnNameFromIndex(2), outputCurrentRow, "シート名"));

                int dataColOutputIdx = 0;
                for (int col = startColumnNum; col <= endColumnNum; col++)
                {
                    headerRow.Append(CreateTextCell(GetColumnNameFromIndex(3 + dataColOutputIdx), outputCurrentRow, $"Data {dataColOutputIdx + 1}"));
                    dataColOutputIdx++;
                }
                sheetData.Append(headerRow);
                outputCurrentRow++;

                string[] fileExtensions = { "*.xlsx", "*.xlsm" };
                List<string> filesToProcess = new List<string>();
                SearchOption searchOption = includeSubdirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly; // Use searchOption

                foreach (string ext in fileExtensions)
                {
                    try
                    {
                        filesToProcess.AddRange(Directory.GetFiles(folderPath, ext, searchOption));
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        logger.Report($"警告: フォルダ '{folderPath}' またはそのサブフォルダへのアクセスが拒否されました。詳細: {ex.Message}");
                    }
                    catch (DirectoryNotFoundException ex)
                    {
                         logger.Report($"警告: フォルダ '{folderPath}' が見つかりません。詳細: {ex.Message}");
                         statusUpdater.Report("エラー: 対象フォルダなし");
                         return; // Stop processing if the base directory is not found
                    }
                }
                
                // Remove duplicates that might occur if folderPath itself matches a sub-folder name in some complex scenarios
                // or if GetFiles returns overlapping results based on how OS handles it (though rare for distinct extensions).
                filesToProcess = filesToProcess.Distinct(StringComparer.OrdinalIgnoreCase).ToList();


                if (!filesToProcess.Any())
                {
                    logger.Report("対象フォルダに処理対象のExcelファイルが見つかりませんでした。");
                    statusUpdater.Report("対象ファイルなし");
                    return;
                }

                foreach (string filePath in filesToProcess)
                {
                    if (Path.GetFileName(filePath).Equals(outputFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue; 
                    }

                    string fileNameOnly = Path.GetFileName(filePath);
                    logger.Report($"処理中のファイル: {filePath}"); // Log full path for clarity
                    statusUpdater.Report($"処理中: {fileNameOnly}");
                    
                    try
                    {
                        using (SpreadsheetDocument sourceDoc = SpreadsheetDocument.Open(filePath, false))
                        {
                            WorkbookPart sourceWorkbookPart = sourceDoc.WorkbookPart;
                            if (sourceWorkbookPart?.Workbook?.Sheets == null)
                            {
                                logger.Report($"警告: {fileNameOnly} のワークブック構造を読み取れませんでした。スキップします。");
                                continue;
                            }

                            foreach (Sheet sourceSheet in sourceWorkbookPart.Workbook.Sheets.Elements<Sheet>())
                            {
                                if (sourceSheet?.Name == null || sourceSheet.Id?.Value == null) continue;
                                string sheetName = sourceSheet.Name;
                                logger.Report($"  処理中のシート: {sheetName}");

                                WorksheetPart sourceWorksheetPart = (WorksheetPart)sourceWorkbookPart.GetPartById(sourceSheet.Id.Value);
                                
                                uint currentSourceRow = startRowIndex; 

                                while (true)
                                {
                                    string firstCellInRowRef = startColumnName + currentSourceRow;
                                    string firstCellValue = GetCellValue(sourceWorkbookPart, sourceWorksheetPart, firstCellInRowRef);

                                    if (string.IsNullOrEmpty(firstCellValue))
                                    {
                                        logger.Report($"    シート '{sheetName}' の行 {currentSourceRow} ({startColumnName}{currentSourceRow}) で空のセルを検出。このシートの処理を終了します。");
                                        break; 
                                    }

                                    Row dataRow = new Row() { RowIndex = outputCurrentRow };
                                    dataRow.Append(CreateTextCell(GetColumnNameFromIndex(0), outputCurrentRow, filePath));
                                    dataRow.Append(CreateTextCell(GetColumnNameFromIndex(1), outputCurrentRow, fileNameOnly));
                                    dataRow.Append(CreateTextCell(GetColumnNameFromIndex(2), outputCurrentRow, sheetName));

                                    int outputDataColLetterIdx = 3; 
                                    for (int colIdx = startColumnNum; colIdx <= endColumnNum; colIdx++)
                                    {
                                        string currentSourceColName = GetColumnNameFromIndex(colIdx);
                                        string cellRef = currentSourceColName + currentSourceRow;
                                        string cellValueStr = GetCellValue(sourceWorkbookPart, sourceWorksheetPart, cellRef);
                                        dataRow.Append(CreateTextCell(GetColumnNameFromIndex(outputDataColLetterIdx++), outputCurrentRow, cellValueStr ?? string.Empty));
                                    }

                                    sheetData.Append(dataRow);
                                    outputCurrentRow++;
                                    currentSourceRow++;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                         logger.Report($"ファイル処理エラー {fileNameOnly} ({filePath}): {ex.Message}。このファイルをスキップします。");
                    }
                }
                workbookPart.Workbook.Save();
                logger.Report($"処理完了。出力ファイル: {outputFilePath}");
            }
        }

        private static Cell CreateTextCell(string columnLetter, uint rowIndex, string text)
        {
            return new Cell(new CellValue(text))
            {
                CellReference = columnLetter + rowIndex,
                DataType = CellValues.String 
            };
        }

        private static string GetCellValue(WorkbookPart workbookPart, WorksheetPart worksheetPart, string cellAddress)
        {
            Cell theCell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellAddress);

            if (theCell == null || theCell.CellValue == null)
            {
                return null; 
            }

            string value = theCell.CellValue.InnerText;
            if (theCell.DataType != null && theCell.DataType.Value == CellValues.SharedString)
            {
                SharedStringTablePart sstPart = workbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sstPart != null && int.TryParse(value, out int sstId))
                {
                    if (sstId >= 0 && sstId < sstPart.SharedStringTable.ChildElements.Count)
                    {
                         return sstPart.SharedStringTable.ChildElements[sstId].InnerText;
                    }
                    return $"#SST_ID_ERR({sstId})"; 
                }
                else if (sstPart == null && theCell.DataType.Value == CellValues.SharedString)
                {
                    return value; 
                }
                return value; 
            }
            return value;
        }
        
        public static int GetColumnIndexFromName(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException(nameof(columnName));
            int index = 0;
            columnName = columnName.ToUpper();
            for (int i = 0; i < columnName.Length; i++)
            {
                index *= 26;
                index += (columnName[i] - 'A' + 1);
            }
            return index - 1; 
        }

        public static string GetColumnNameFromIndex(int columnIndex)
        {
            int dividend = columnIndex + 1; 
            string columnName = String.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }
        
        public static void GetCellRowColumn(string cellReference, out uint rowIndex, out string columnName)
        {
            if (string.IsNullOrEmpty(cellReference)) throw new ArgumentNullException(nameof(cellReference));
            cellReference = cellReference.ToUpper();

            Match match = Regex.Match(cellReference, @"([A-Z]+)(\d+)");
            if (!match.Success) throw new ArgumentException("Invalid cell reference format.", nameof(cellReference));

            columnName = match.Groups[1].Value;
            rowIndex = uint.Parse(match.Groups[2].Value);
            if (rowIndex == 0) throw new ArgumentException("Row index cannot be 0.", nameof(cellReference));
        }
        
        public static bool IsValidCellReference(string cellRef)
        {
            if (string.IsNullOrWhiteSpace(cellRef)) return false;
            return Regex.IsMatch(cellRef.ToUpper(), @"^[A-Z]+[1-9][0-9]*$");
        }
    }
}