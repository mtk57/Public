using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
// ★ 以下のusingディレクティブを追加
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace ER2SpreadTool
{


    public partial class MainForm : Form
    {
        private readonly string settingsFilePath;

        public MainForm()
        {
            InitializeComponent();
            settingsFilePath = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "er2spreadtool_settings.json");
            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            LoadSettings();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        private void LoadSettings()
        {
            if (!File.Exists(settingsFilePath))
            {
                txtResults.AppendText("設定ファイルが見つかりませんでした。初回起動または設定ファイルが削除されています。\n");
                return;
            }

            try
            {
                using (FileStream fs = new FileStream(settingsFilePath, FileMode.Open, FileAccess.Read))
                {
                    if (fs.Length == 0) // 空のファイルの場合
                    {
                        txtResults.AppendText("設定ファイルは空です。\n");
                        return;
                    }
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    AppSettings settings = (AppSettings)serializer.ReadObject(fs);

                    if (settings != null)
                    {
                        txtFilePath.Text = settings.LastFilePath;
                        txtSheetName.Text = settings.LastSheetName;
                        txtResults.AppendText("前回終了時の設定を読み込みました。\n");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの読み込み中にエラーが発生しました:\n{ex.Message}", "設定読み込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtResults.AppendText($"設定ファイルの読み込みエラー: {ex.Message}\n");
            }
        }

        private void SaveSettings()
        {
            AppSettings settings = new AppSettings
            {
                LastFilePath = txtFilePath.Text,
                LastSheetName = txtSheetName.Text
            };

            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    serializer.WriteObject(ms, settings);
                    byte[] jsonBytes = ms.ToArray();
                    // インデントされたJSONとして書き出すために、一度文字列に変換して手動でインデント風にするか、
                    // またはDataContractJsonSerializerSettingsでインデントオプションを探す（標準では単純なインデントは難しい）
                    // ここでは単純な書き出しとします。
                    File.WriteAllBytes(settingsFilePath, jsonBytes);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの保存中にエラーが発生しました:\n{ex.Message}", "設定保存エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                // 必要であればログに書き出すなどの処理を追加
            }
        }

        // btnBrowse_Click メソッドを少し変更して、読み込んだパスを利用
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excelファイル (*.xlsx)|*.xlsx";
                openFileDialog.Title = "処理対象のExcelファイルを選択してください";

                // 前回値があればそれを初期値とする
                if (!string.IsNullOrWhiteSpace(txtFilePath.Text))
                {
                    if (File.Exists(txtFilePath.Text))
                    {
                        openFileDialog.InitialDirectory = Path.GetDirectoryName(txtFilePath.Text);
                        openFileDialog.FileName = Path.GetFileName(txtFilePath.Text);
                    }
                    else if (Directory.Exists(txtFilePath.Text)) // もしパスがディレクトリなら
                    {
                        openFileDialog.InitialDirectory = txtFilePath.Text;
                    }
                }


                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = openFileDialog.FileName;
                }
            }
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFilePath.Text))
            {
                MessageBox.Show("Excelファイルパスを指定してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(txtSheetName.Text))
            {
                MessageBox.Show("シート名を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!File.Exists(txtFilePath.Text))
            {
                MessageBox.Show("指定されたファイルが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                lblStatus.Text = "処理中...";
                // txtResults.Clear(); // ログのクリアタイミングは要望に応じて調整
                Application.DoEvents();

                var results = ProcessExcelFile(txtFilePath.Text, txtSheetName.Text);

                if (results.Any())
                {
                    txtResults.AppendText($"\n抽出結果:\n{FormatResultsForDisplay(results)}");
                    
                    // TSVファイル出力（入力ファイルを変更しない）
                    string outputFilePath = CreateTsvOutput(txtFilePath.Text, results);
                    
                    lblStatus.Text = $"処理完了 - {results.Count}件のテーブル情報を抽出しました。";
                    MessageBox.Show($"処理が完了しました。\nTSVファイルを出力しました:\n{outputFilePath}", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    txtResults.AppendText("\n処理対象の図形が見つかりませんでした。\nデバッグ情報を確認してください。");
                    lblStatus.Text = "処理対象の図形が見つかりませんでした。";
                    MessageBox.Show("処理対象の図形が見つかりませんでした。ログを確認してください。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = "処理エラー";
                txtResults.AppendText($"\nエラーが発生しました:\n{ex.ToString()}\n");
                MessageBox.Show($"処理中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<TableInfo> ProcessExcelFile(string filePath, string sheetName)
        {
            var results = new List<TableInfo>();
            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null)
                {
                    txtResults.AppendText("WorkbookPart が見つかりません。\n");
                    return results;
                }

                var sheet = workbookPart.Workbook.Descendants<Sheet>()
                    .FirstOrDefault(s => s.Name != null && s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                if (sheet == null) throw new ArgumentException($"シート '{sheetName}' が見つかりません。");

                var worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                if (worksheetPart == null)
                {
                    txtResults.AppendText($"WorksheetPart が見つかりません (シートID: {sheet.Id.Value})。\n");
                    return results;
                }
                txtResults.AppendText($"シート '{sheetName}' (ID: {sheet.Id.Value}) を処理中...\n");

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart == null)
                {
                    txtResults.AppendText("DrawingsPart が見つかりません。図形が存在しない可能性があります。\n");
                    return results;
                }

                var worksheetDrawing = drawingsPart.WorksheetDrawing;
                if (worksheetDrawing == null)
                {
                    txtResults.AppendText("WorksheetDrawing が見つかりません。\n");
                    return results;
                }
                txtResults.AppendText("WorksheetDrawing を取得しました。\n");
                LogAllShapes(worksheetDrawing);

                foreach (var anchor in worksheetDrawing.Elements<TwoCellAnchor>())
                {
                    // Iterate through DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape (xdr:grpSp) elements
                    foreach (var xdrGroupShape in anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape>())
                    {
                        txtResults.AppendText("Spreadsheet.GroupShape (xdr:grpSp) をTwoCellAnchor内で発見。\n");
                        var tableInfo = ExtractTableInfoFromSpreadsheetGroup(xdrGroupShape);
                        if (tableInfo != null)
                        {
                            if (!results.Any(r => r.TableName.Equals(tableInfo.TableName, StringComparison.OrdinalIgnoreCase)))
                            {
                                results.Add(tableInfo);
                                txtResults.AppendText($"  => グループからテーブル抽出成功: {tableInfo.TableName} (列数: {tableInfo.Columns.Count})\n");
                            }
                            else
                            {
                                txtResults.AppendText($"  => 重複テーブル名のためスキップ: {tableInfo.TableName}\n");
                            }
                        }
                        else
                        {
                            txtResults.AppendText("  => Spreadsheet.GroupShapeからテーブル情報抽出失敗。\n");
                        }
                    }
                }
            }
            return results;
        }

        private TableInfo ExtractTableInfoFromSpreadsheetGroup(DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape xdrGroupShape)
        {
            txtResults.AppendText("  ExtractTableInfoFromSpreadsheetGroup (xdr:grpSp) 開始\n");

            // Scenario 1: The xdr:grpSp contains a nested a:grpSp (Drawing.GroupShape)
            var nestedDrawingGroup = xdrGroupShape.GetFirstChild<DocumentFormat.OpenXml.Drawing.GroupShape>();
            if (nestedDrawingGroup != null)
            {
                txtResults.AppendText("    ネストされた Drawing.GroupShape (a:grpSp) を発見。これを処理します。\n");
                return ExtractTableInfoFromDrawingGroupShape(nestedDrawingGroup);
            }

            // Scenario 2: The xdr:grpSp directly contains two xdr:sp (Spreadsheet.Shape) elements
            var xdrShapesInGroup = xdrGroupShape.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().ToList();
            txtResults.AppendText($"    xdr:grpSp内のSpreadsheet.Shape (xdr:sp) 数: {xdrShapesInGroup.Count}\n");

            if (xdrShapesInGroup.Count < 2) // Needs at least one for table name, one for columns
            {
                txtResults.AppendText($"    xdr:grpSp内に少なくとも2つのxdr:spが必要です (テーブル名用1、カラム名用1以上)。実際は{xdrShapesInGroup.Count}個。\n");
                return null;
            }

            var textDataFromXdrShapes = new List<ShapeTextParseResult>();
            for (int i = 0; i < xdrShapesInGroup.Count; i++)
            {
                var xdrShape = xdrShapesInGroup[i];
                txtResults.AppendText($"    xdr:sp {i + 1} のテキスト抽出試行...\n");
                string textContent = string.Empty;

                var spreadsheetTextBody = xdrShape.TextBody; // This is DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody
                if (spreadsheetTextBody != null)
                {
                    txtResults.AppendText("      Spreadsheet.TextBody (xdr:txBody) を発見。\n");
                    textContent = ExtractTextFromSpreadsheetTextBody(spreadsheetTextBody);
                }
                else
                {
                    txtResults.AppendText("      Spreadsheet.TextBody (xdr:txBody) が見つかりません。\n");
                }

                string[] lines = new string[0];
                if (!string.IsNullOrWhiteSpace(textContent))
                {
                    lines = textContent.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                       .Select(l => l.Trim())
                                       .Where(l => !string.IsNullOrEmpty(l))
                                       .ToArray();
                    txtResults.AppendText($"      抽出テキスト: '{textContent.Replace("\n", "\\n")}' (有効行数: {lines.Length})\n");
                }
                textDataFromXdrShapes.Add(new ShapeTextParseResult { OriginalText = textContent, Lines = lines });
            }
            return CreateTableInfoFromTextData(textDataFromXdrShapes, "xdr:spベース");
        }

        private TableInfo ExtractTableInfoFromDrawingGroupShape(DocumentFormat.OpenXml.Drawing.GroupShape drawingGroupShape)
        {
            txtResults.AppendText("  ExtractTableInfoFromDrawingGroupShape (a:grpSp) 開始\n");
            var shapesInDrawingGroup = drawingGroupShape.Elements<DocumentFormat.OpenXml.Drawing.Shape>().ToList(); // These are a:sp
            txtResults.AppendText($"    a:grpSp内のDrawing.Shape (a:sp) 数: {shapesInDrawingGroup.Count}\n");

            if (shapesInDrawingGroup.Count < 2) // Needs at least one for table name, one for columns
            {
                txtResults.AppendText($"    a:grpSp内に少なくとも2つのa:spが必要です (テーブル名用1、カラム名用1以上)。実際は{shapesInDrawingGroup.Count}個。\n");
                return null;
            }

            var textDataFromDrawingShapes = new List<ShapeTextParseResult>();
            for (int i = 0; i < shapesInDrawingGroup.Count; i++)
            {
                var drawingShape = shapesInDrawingGroup[i]; // This is a:sp
                txtResults.AppendText($"    a:sp {i + 1} のテキスト抽出試行...\n");
                var drawingTextBody = drawingShape.GetFirstChild<DocumentFormat.OpenXml.Drawing.TextBody>(); // a:sp -> a:txBody
                string textContent = string.Empty;
                string[] lines = new string[0];

                if (drawingTextBody != null)
                {
                    textContent = ExtractTextFromDrawingTextBody(drawingTextBody);
                    if (!string.IsNullOrWhiteSpace(textContent))
                    {
                        lines = textContent.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                           .Select(l => l.Trim())
                                           .Where(l => !string.IsNullOrEmpty(l))
                                           .ToArray();
                        txtResults.AppendText($"      抽出テキスト: '{textContent.Replace("\n", "\\n")}' (有効行数: {lines.Length})\n");
                    }
                }
                else
                {
                    txtResults.AppendText($"      a:sp {i + 1} に Drawing.TextBody (a:txBody) が見つかりません。\n");
                }
                textDataFromDrawingShapes.Add(new ShapeTextParseResult { OriginalText = textContent, Lines = lines });
            }
            return CreateTableInfoFromTextData(textDataFromDrawingShapes, "a:spベース");
        }

        private TableInfo CreateTableInfoFromTextData(List<ShapeTextParseResult> textDataList, string sourceDescription)
        {
            txtResults.AppendText($"    {sourceDescription}: CreateTableInfoFromTextData 開始。ShapeTextParseResult 数: {textDataList.Count}\n");

            if (textDataList.Count < 2)
            {
                txtResults.AppendText($"    {sourceDescription}: 少なくとも2つのテキスト情報が必要です (テーブル名用1つ、カラム名用1つ以上)。実際は{textDataList.Count}個。\n");
                return null;
            }

            // Identify potential table name shapes (those with exactly one line of non-empty text)
            List<ShapeTextParseResult> potentialTableShapes = textDataList
                .Where(s => s.Lines != null && s.Lines.Length == 1 && !string.IsNullOrWhiteSpace(s.Lines[0]))
                .ToList();

            ShapeTextParseResult tableShapeData = null;
            List<string> columnLinesAggregated = new List<string>();

            if (potentialTableShapes.Count == 1)
            {
                tableShapeData = potentialTableShapes[0];
                txtResults.AppendText($"    {sourceDescription}: テーブル名候補のシェイプを発見 (元テキスト: '{tableShapeData.OriginalText.Replace("\n", "\\n")}')\n");

                // All other shapes contribute to columns
                foreach (var shapeData in textDataList)
                {
                    if (shapeData != tableShapeData) // If it's not the table name shape
                    {
                        if (shapeData.Lines != null && shapeData.Lines.Length > 0)
                        {
                            columnLinesAggregated.AddRange(shapeData.Lines);
                            txtResults.AppendText($"    {sourceDescription}: カラム候補のシェイプから行を追加 (元テキスト: '{shapeData.OriginalText.Replace("\n", "\\n")}', 追加行数: {shapeData.Lines.Length})\n");
                        }
                        else
                        {
                            txtResults.AppendText($"    {sourceDescription}: カラム候補シェイプに有効な行なし、または空テキスト (元テキスト: '{shapeData.OriginalText.Replace("\n", "\\n")}')\n");
                        }
                    }
                }
            }
            else
            {
                txtResults.AppendText($"    {sourceDescription}: テーブル名となるシェイプ (テキスト1行のみ) が1つである必要がありますが、{potentialTableShapes.Count}個見つかりました。\n");
                txtResults.AppendText($"      調査対象シェイプ数: {textDataList.Count}\n");
                for (int i = 0; i < textDataList.Count; i++)
                {
                    var currentShape = textDataList[i];
                    txtResults.AppendText($"        Shape {i + 1} 有効行数: {(currentShape.Lines?.Length ?? 0)} (元テキスト: '{currentShape.OriginalText.Replace("\n", "\\n")}')\n");
                }
                return null;
            }

            // Validate that a table name was indeed found and is not empty.
            if (tableShapeData == null) // Should ideally not happen if potentialTableShapes.Count == 1
            {
                txtResults.AppendText($"    {sourceDescription}: テーブル名シェイプの特定に失敗しました (内部エラー)。\n");
                return null;
            }
            string tableName = tableShapeData.Lines[0].Trim();


            // Validate that column data was found.
            if (columnLinesAggregated.Count == 0)
            {
                txtResults.AppendText($"    {sourceDescription}: 抽出されたカラムリストが空です (テーブル名: '{tableName}')。\n");
                return null;
            }

            // Process and deduplicate column names
            List<string> finalColumns = columnLinesAggregated
                                        .Where(l => !string.IsNullOrWhiteSpace(l)) // Ensure no empty/whitespace lines become columns
                                        .Select(l => l.Trim())                     // Trim whitespace from each column name
                                        .Distinct(StringComparer.OrdinalIgnoreCase) // Deduplicate column names (case-insensitive)
                                        .ToList();

            if (finalColumns.Count == 0) // After trimming and distinct, it might become empty
            {
                txtResults.AppendText($"    {sourceDescription}: 有効なカラム名が抽出・処理後、空になりました (テーブル名: '{tableName}')。\n");
                return null;
            }

            txtResults.AppendText($"    {sourceDescription}: テーブル名として確定: '{tableName}'\n");
            txtResults.AppendText($"    {sourceDescription}: カラム名として確定 ({finalColumns.Count}件): [{string.Join(", ", finalColumns)}]\n");

            return new TableInfo { TableName = tableName, Columns = finalColumns };
        }

        private string ExtractTextFromDrawingTextBody(DocumentFormat.OpenXml.Drawing.TextBody drawingTextBody) // For a:txBody
        {
            var textParts = new List<string>();
            foreach (var paragraph in drawingTextBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>()) // a:p
            {
                string paragraphText = "";
                foreach (var run in paragraph.Elements<DocumentFormat.OpenXml.Drawing.Run>()) // a:r
                {
                    foreach (var text in run.Elements<DocumentFormat.OpenXml.Drawing.Text>()) // a:t
                    {
                        if (text.Text != null)
                        {
                            paragraphText += text.Text;
                        }
                    }
                }
                textParts.Add(paragraphText);
            }
            return string.Join("\n", textParts);
        }

        private string ExtractTextFromSpreadsheetTextBody(DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody spreadsheetTextBody) // For xdr:txBody
        {
            var textParts = new List<string>();
            // xdr:txBody also contains a:p elements
            foreach (var paragraph in spreadsheetTextBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                string paragraphText = "";
                foreach (var run in paragraph.Elements<DocumentFormat.OpenXml.Drawing.Run>())
                {
                    foreach (var text in run.Elements<DocumentFormat.OpenXml.Drawing.Text>())
                    {
                        if (text.Text != null)
                        {
                            paragraphText += text.Text;
                        }
                    }
                }
                textParts.Add(paragraphText);
            }
            return string.Join("\n", textParts);
        }

        private void LogAllShapes(WorksheetDrawing worksheetDrawing)
        {
            txtResults.AppendText("=== 図形構造の調査 ===\n");
            var childElements = worksheetDrawing.ChildElements.ToList();
            txtResults.AppendText($"WorksheetDrawing の直下の子要素数: {childElements.Count}\n");

            foreach (var element in childElements)
            {
                txtResults.AppendText($"要素タイプ: {element.GetType().FullName}\n"); // Use FullName for clarity
                if (element is TwoCellAnchor anchor)
                {
                    LogTwoCellAnchor(anchor);
                }
            }
            txtResults.AppendText("=== 調査完了 ===\n\n");
        }

        private void LogTwoCellAnchor(TwoCellAnchor anchor)
        {
            txtResults.AppendText("  TwoCellAnchor 内容:\n");
            foreach (var child in anchor.ChildElements)
            {
                txtResults.AppendText($"    子要素タイプ: {child.GetType().FullName}\n");
                if (child is DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape xdrShape)
                {
                    txtResults.AppendText("      詳細: Spreadsheet.Shape (xdr:sp)\n");
                    var spTextBody = xdrShape.TextBody;
                    if (spTextBody != null)
                    {
                        txtResults.AppendText($"        xdr:sp Text: '{ExtractTextFromSpreadsheetTextBody(spTextBody).Replace("\n", "\\n")}'\n");
                    }
                }
                else if (child is DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape xdrGroupShape)
                {
                    txtResults.AppendText("      詳細: Spreadsheet.GroupShape (xdr:grpSp)\n");
                    foreach (var innerElement in xdrGroupShape.Elements<OpenXmlCompositeElement>())
                    {
                        txtResults.AppendText($"        xdr:grpSp 内の子要素: {innerElement.GetType().FullName}\n");
                        if (innerElement is DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape innerXdrShape)
                        {
                            var innerSpTextBody = innerXdrShape.TextBody;
                            if (innerSpTextBody != null)
                            {
                                txtResults.AppendText($"          xdr:sp (内) Text: '{ExtractTextFromSpreadsheetTextBody(innerSpTextBody).Replace("\n", "\\n")}'\n");
                            }
                        }
                        else if (innerElement is DocumentFormat.OpenXml.Drawing.GroupShape innerDrawingGroupShape)
                        { // a:grpSp
                            txtResults.AppendText($"          Drawing.GroupShape (a:grpSp) (内)\n");
                        }
                    }
                }
            }
        }

        private string FormatResultsForDisplay(List<TableInfo> results)
        {
            var output = new List<string>();
            output.Add("--- 抽出テーブル一覧 ---");
            foreach (var table in results)
            {
                output.Add($"テーブル名: {table.TableName}");
                output.Add("  カラム名:");
                foreach (var column in table.Columns)
                {
                    output.Add($"    - {column}");
                }
                output.Add("");
            }
            return string.Join("\n", output);
        }

        // TSVファイル出力メソッド（新規追加）
        private string CreateTsvOutput(string inputFilePath, List<TableInfo> results)
        {
            // 入力ファイルと同じフォルダにTSVファイルを作成
            string inputDirectory = Path.GetDirectoryName(inputFilePath);
            
            // 現在日時をYYYYMMDD_hhmmssfff形式でフォーマット
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
            string outputFileName = $"{timestamp}.tsv";
            string outputFilePath = Path.Combine(inputDirectory, outputFileName);

            try
            {
                using (var writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
                {
                    // ヘッダー行を出力
                    writer.WriteLine("#\tnum\tテーブル名\tカラム名");

                    // データ行を出力
                    int overallCounter = 1;
                    foreach (var table in results)
                    {
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            writer.WriteLine($"{overallCounter}\t{i + 1}\t{table.TableName}\t{table.Columns[i]}");
                            overallCounter++;
                        }
                    }
                }

                txtResults.AppendText($"TSVファイルを出力しました: {outputFilePath}\n");
                return outputFilePath;
            }
            catch (Exception ex)
            {
                string errorMessage = $"TSVファイルの出力中にエラーが発生しました: {ex.Message}";
                txtResults.AppendText($"{errorMessage}\n");
                MessageBox.Show(errorMessage, "出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }
    }

    // 設定を保持するクラス (DataContract属性を使用)
    [DataContract]
    internal class AppSettings
    {
        [DataMember]
        public string LastFilePath { get; set; }

        [DataMember]
        public string LastSheetName { get; set; }
    }

    internal class TableInfo
    {
        public string TableName { get; set; }
        public List<string> Columns { get; set; } = new List<string>();
    }

    internal class ShapeTextParseResult
    {
        public string OriginalText { get; set; }
        public string[] Lines { get; set; }
    }
}