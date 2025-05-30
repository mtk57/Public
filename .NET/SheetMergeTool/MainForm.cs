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
// ★ 以下のusingディレクティブを追加 (設定ファイル対応)
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace SheetMergeTool
{
    public partial class MainForm : Form
    {
        // ★ 設定ファイルのパスを保持するフィールド
        private readonly string settingsFilePath;

        public MainForm()
        {
            InitializeComponent();

            // ★ 設定ファイルのパスを初期化 (アプリケーションと同じディレクトリ)
            settingsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sheetmergetool_settings.json");

            // ★ イベントハンドラの登録 (設定ファイルのロード/セーブ)
            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;

            lblStatus.Text = "準備完了"; // Initial status

            // ★ ドラッグアンドドロップ機能の有効化とイベントハンドラの設定 (フォルダパス用)
            this.txtDirPath.AllowDrop = true;
            this.txtDirPath.DragEnter += new DragEventHandler(txtDirPath_DragEnter);
            this.txtDirPath.DragDrop += new DragEventHandler(txtDirPath_DragDrop);
        }

        // ★ MainForm_Load イベントハンドラ (設定読み込み)
        private void MainForm_Load(object sender, EventArgs e)
        {
            LoadSettings();
        }

        // ★ MainForm_FormClosing イベントハンドラ (設定保存)
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        // ★ 設定読み込み処理
        private void LoadSettings()
        {
            if (!File.Exists(settingsFilePath))
            {
                if (!txtResults.IsDisposed) // Check if txtResults still exists
                {
                    txtResults.AppendText("設定ファイルが見つかりませんでした。初回起動または設定ファイルが削除されています。\n");
                }
                return;
            }

            try
            {
                using (FileStream fs = new FileStream(settingsFilePath, FileMode.Open, FileAccess.Read))
                {
                    if (fs.Length == 0) // 空のファイルの場合
                    {
                        if (!txtResults.IsDisposed)
                        {
                            txtResults.AppendText("設定ファイルは空です。\n");
                        }
                        return;
                    }
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(SheetMergeToolSettings));
                    SheetMergeToolSettings settings = (SheetMergeToolSettings)serializer.ReadObject(fs);

                    if (settings != null)
                    {
                        txtDirPath.Text = settings.LastFolderPath;
                        txtStartCell.Text = settings.LastStartCell;
                        txtEndCell.Text = settings.LastEndCell;
                        chkEnableSubDir.Checked = settings.LastIncludeSubdirectories;
                        if (!txtResults.IsDisposed)
                        {
                            txtResults.AppendText("前回終了時の設定を読み込みました。\n");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの読み込み中にエラーが発生しました:\n{ex.Message}", "設定読み込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (!txtResults.IsDisposed)
                {
                    txtResults.AppendText($"設定ファイルの読み込みエラー: {ex.Message}\n");
                }
            }
        }

        // ★ 設定保存処理
        private void SaveSettings()
        {
            SheetMergeToolSettings settings = new SheetMergeToolSettings
            {
                LastFolderPath = txtDirPath.Text,
                LastStartCell = txtStartCell.Text,
                LastEndCell = txtEndCell.Text,
                LastIncludeSubdirectories = chkEnableSubDir.Checked
            };

            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(SheetMergeToolSettings));
                    serializer.WriteObject(ms, settings);
                    File.WriteAllBytes(settingsFilePath, ms.ToArray());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの保存中にエラーが発生しました:\n{ex.Message}", "設定保存エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ★ txtDirPath の DragEnter イベントハンドラ (フォルダパス用)
        private void txtDirPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0 && Directory.Exists(files[0]))
                {
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        // ★ txtDirPath の DragDrop イベントハンドラ (フォルダパス用)
        private void txtDirPath_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (files != null && files.Length > 0)
            {
                string selectedPath = files[0];
                if (Directory.Exists(selectedPath))
                {
                    this.txtDirPath.Text = selectedPath;
                }
                else
                {
                    MessageBox.Show("ドロップされたアイテムは有効なフォルダではありません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Excelファイルが含まれるフォルダを選択してください";
                // 前回選択したフォルダパスを初期表示 (設定ファイルから読み込まれた値を利用)
                if (!string.IsNullOrWhiteSpace(txtDirPath.Text) && Directory.Exists(txtDirPath.Text))
                {
                    dialog.SelectedPath = txtDirPath.Text;
                }
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
            bool includeSubdirectories = chkEnableSubDir.Checked;

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
            if (!txtResults.IsDisposed)
            {
                 txtResults.Clear(); // 既存のログをクリア
            }
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
                    MergeExcelFilesLogic(folderPath, startCellRef, endCellRef, includeSubdirectories, progressLog, progressStatus)
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
            chkEnableSubDir.Enabled = enabled;
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
                SearchOption searchOption = includeSubdirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

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
                         return; 
                    }
                }
                
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
                    logger.Report($"処理中のファイル: {filePath}"); 
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
                                string sheetName = sourceSheet.Name.Value; 
                                logger.Report($"  処理中のシート: {sheetName}");

                                WorksheetPart sourceWorksheetPart = (WorksheetPart)sourceWorkbookPart.GetPartById(sourceSheet.Id.Value);
                                
                                uint currentSourceRow = startRowIndex; 

                                while (true)
                                {
                                    // ▼▼▼ 変更点 ▼▼▼
                                    bool isRowCompletelyEmpty = true;
                                    List<string> currentRowCellValues = new List<string>();

                                    // 指定された列範囲のセル値をチェック
                                    for (int colIdx = startColumnNum; colIdx <= endColumnNum; colIdx++)
                                    {
                                        string currentSourceColName = GetColumnNameFromIndex(colIdx);
                                        string cellRef = currentSourceColName + currentSourceRow;
                                        string cellValueStr = GetCellValue(sourceWorkbookPart, sourceWorksheetPart, cellRef);
                                        
                                        currentRowCellValues.Add(cellValueStr ?? string.Empty); // 値をリストに追加（nullの場合は空文字）

                                        if (!string.IsNullOrEmpty(cellValueStr))
                                        {
                                            isRowCompletelyEmpty = false; // 1つでも値があれば、行は空ではない
                                        }
                                    }

                                    if (isRowCompletelyEmpty)
                                    {
                                        logger.Report($"    シート '{sheetName}' の行 {currentSourceRow} (列 {startColumnName} から {endColumnName}) の全セルが空のため、このシートの処理を終了します。");
                                        break; 
                                    }
                                    // ▲▲▲ 変更点 ▲▲▲

                                    Row dataRow = new Row() { RowIndex = outputCurrentRow };
                                    dataRow.Append(CreateTextCell(GetColumnNameFromIndex(0), outputCurrentRow, filePath));
                                    dataRow.Append(CreateTextCell(GetColumnNameFromIndex(1), outputCurrentRow, fileNameOnly));
                                    dataRow.Append(CreateTextCell(GetColumnNameFromIndex(2), outputCurrentRow, sheetName));

                                    int outputDataColLetterIdx = 3; 
                                    // ▼▼▼ 変更点 ▼▼▼
                                    // 収集したセル値を使って出力行を生成
                                    foreach (string cellValue in currentRowCellValues)
                                    {
                                        dataRow.Append(CreateTextCell(GetColumnNameFromIndex(outputDataColLetterIdx++), outputCurrentRow, cellValue));
                                    }
                                    // ▲▲▲ 変更点 ▲▲▲
                                    
                                    sheetData.Append(dataRow);
                                    outputCurrentRow++;
                                    currentSourceRow++;
                                }
                            }
                        }
                    }
                    catch (OpenXmlPackageException oxpe) 
                    {
                         logger.Report($"ファイル処理エラー (OpenXML) {fileNameOnly} ({filePath}): {oxpe.Message}。このファイルをスキップします。破損しているか、パスワードで保護されている可能性があります。");
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
                DataType = new EnumValue<CellValues>(CellValues.String) 
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
                if (sstPart?.SharedStringTable != null && int.TryParse(value, out int sstId)) 
                {
                    if (sstId >= 0 && sstId < sstPart.SharedStringTable.ChildElements.Count)
                    {
                         return sstPart.SharedStringTable.ChildElements[sstId].InnerText;
                    }
                    return $"#SST_ID_OOR({sstId})"; 
                }
                return value; 
            }
            return value;
        }
        
        public static int GetColumnIndexFromName(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException(nameof(columnName));
            int index = 0;
            columnName = columnName.ToUpperInvariant(); 
            for (int i = 0; i < columnName.Length; i++)
            {
                if (columnName[i] < 'A' || columnName[i] > 'Z') throw new ArgumentException("Invalid character in column name.", nameof(columnName));
                index *= 26;
                index += (columnName[i] - 'A' + 1);
            }
            return index - 1; 
        }

        public static string GetColumnNameFromIndex(int columnIndex)
        {
            if (columnIndex < 0) throw new ArgumentOutOfRangeException(nameof(columnIndex), "Column index must be non-negative.");
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
            cellReference = cellReference.ToUpperInvariant(); 

            Match match = Regex.Match(cellReference, @"^([A-Z]+)([1-9][0-9]*)$"); 
            if (!match.Success) throw new ArgumentException("Invalid cell reference format.", nameof(cellReference));

            columnName = match.Groups[1].Value;
            rowIndex = uint.Parse(match.Groups[2].Value);
        }
        
        public static bool IsValidCellReference(string cellRef)
        {
            if (string.IsNullOrWhiteSpace(cellRef)) return false;
            return Regex.IsMatch(cellRef.ToUpperInvariant(), @"^[A-Z]+[1-9][0-9]*$");
        }
    }

    [DataContract]
    internal class SheetMergeToolSettings
    {
        [DataMember]
        public string LastFolderPath { get; set; }

        [DataMember]
        public string LastStartCell { get; set; }

        [DataMember]
        public string LastEndCell { get; set; }

        [DataMember]
        public bool LastIncludeSubdirectories { get; set; }
    }
}