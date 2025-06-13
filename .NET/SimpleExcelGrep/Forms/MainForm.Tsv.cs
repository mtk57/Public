using SimpleExcelGrep.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace SimpleExcelGrep.Forms
{
    /// <summary>
    /// メインフォーム (TSV入出力関連の処理部分)
    /// </summary>
    public partial class MainForm
    {
        /// <summary>
        /// TSV読み込みボタンクリック時の処理
        /// </summary>
        private void BtnLoadTsv_Click(object sender, EventArgs e)
        {
            if (_isSearching)
            {
                MessageBox.Show("検索中はTSVファイルを読み込めません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (grdResults.Rows.Count > 0)
            {
                var dialogResult = MessageBox.Show("現在の検索結果はクリアされます。TSVファイルを読み込みますか？",
                    "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.No) return;
            }

            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "TSV files (*.tsv)|*.tsv|All files (*.*)|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    LoadTsvFile(openFileDialog.FileName);
                }
            }
        }

        /// <summary>
        /// 検索結果をTSVファイルに書き出す
        /// </summary>
        private void WriteResultsToTsv(List<SearchResult> results)
        {
            if (results == null || !results.Any()) return;

            try
            {
                string exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string fileName = $"{DateTime.Now:yyyyMMdd_HHmmss}.tsv";
                string filePath = Path.Combine(exePath, fileName);
                
                var sb = new StringBuilder();
                sb.AppendLine("ファイルパス\tファイル名\tシート名\tセル位置\tセルの値");
                foreach (var result in results)
                {
                    sb.AppendLine(string.Join("\t",
                        EscapeTsvField(result.FilePath),
                        EscapeTsvField(Path.GetFileName(result.FilePath)),
                        EscapeTsvField(result.SheetName),
                        EscapeTsvField(result.CellPosition),
                        EscapeTsvField(result.CellValue)
                    ));
                }

                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                UpdateStatus($"TSVファイルに結果を書き出しました: {fileName}");
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"TSVファイル書き出しエラー: {ex.Message}");
                MessageBox.Show($"TSVファイルへの書き出し中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// TSVファイルを読み込み、DataGridViewに表示する
        /// </summary>
        private void LoadTsvFile(string filePath)
        {
            try
            {
                var lines = File.ReadAllLines(filePath, Encoding.UTF8);
                if (lines.Length <= 1)
                {
                    MessageBox.Show("TSVファイルに読み込むデータがありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                _searchResults.Clear();
                // ヘッダー行をスキップ
                for (int i = 1; i < lines.Length; i++)
                {
                    if (string.IsNullOrWhiteSpace(lines[i])) continue;
                    
                    string[] fields = ParseTsvLine(lines[i]);
                    if (fields.Length >= 5)
                    {
                        _searchResults.Add(new SearchResult
                        {
                            FilePath = UnescapeTsvField(fields[0]),
                            SheetName = UnescapeTsvField(fields[2]),
                            CellPosition = UnescapeTsvField(fields[3]),
                            CellValue = UnescapeTsvField(fields[4])
                        });
                    }
                }
                DisplaySearchResults(_searchResults);
                UpdateStatus($"TSVファイルから {_searchResults.Count} 件の結果を読み込みました。");
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"TSVファイル読み込みエラー: {ex.Message}");
                MessageBox.Show($"TSVファイルの読み込み中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus("TSVファイルの読み込みに失敗しました。");
            }
        }
        
        /// <summary>
        /// TSVフィールドのエスケープ処理
        /// </summary>
        private string EscapeTsvField(string field)
        {
            if (string.IsNullOrEmpty(field)) return "";
            if (field.Contains('\t') || field.Contains('\n') || field.Contains('\r') || field.Contains('"'))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }

        /// <summary>
        /// TSVの1行をパースする
        /// </summary>
        private string[] ParseTsvLine(string line)
        {
            var fields = new List<string>();
            var currentField = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            currentField.Append('"');
                            i++;
                        }
                        else inQuotes = false;
                    }
                    else currentField.Append(c);
                }
                else
                {
                    if (c == '"') inQuotes = true;
                    else if (c == '\t')
                    {
                        fields.Add(currentField.ToString());
                        currentField.Clear();
                    }
                    else currentField.Append(c);
                }
            }
            fields.Add(currentField.ToString());
            return fields.ToArray();
        }

        /// <summary>
        /// TSVフィールドのアンエスケープ処理
        /// </summary>
        private string UnescapeTsvField(string field)
        {
            if (string.IsNullOrEmpty(field)) return "";
            if (field.StartsWith("\"") && field.EndsWith("\""))
            {
                return field.Substring(1, field.Length - 2).Replace("\"\"", "\"");
            }
            return field;
        }
    }
}