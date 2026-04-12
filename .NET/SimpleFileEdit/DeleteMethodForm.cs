using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace SimpleFileSearch
{
    public partial class DeleteMethodForm : Form
    {
        private readonly ILanguageProcessor _processor;
        private readonly string _settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DeleteMethodForm.json");

        public string Keyword { get; private set; }
        public bool IsRegexEnabled { get; private set; }
        public bool IsNotEnabled { get; private set; }

        public DeleteMethodForm(ILanguageProcessor processor)
        {
            InitializeComponent();
            _processor = processor;
            btnDeleteStart.Click += BtnDeleteStart_Click;
            btnCloseForm.Click += BtnCloseForm_Click;
            LoadSettings();
        }

        public Func<string, string> CreateTransformer()
        {
            var keyword = Keyword;
            var useRegex = IsRegexEnabled;
            var negateMatch = IsNotEnabled;
            var processor = _processor;
            return content => RemoveMethodsAndImports(content, keyword, useRegex, negateMatch, processor);
        }

        private void BtnDeleteStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtKeywordForMethodSignature.Text))
            {
                return;
            }

            if (chkEnabledRegExForMethodSignature.Checked)
            {
                try
                {
                    Regex.IsMatch(string.Empty, txtKeywordForMethodSignature.Text);
                }
                catch (ArgumentException)
                {
                    MessageBox.Show(this, "正規表現が不正です。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            var confirm = MessageBox.Show(this,
                "メソッド、import削除を実行します。よろしいですか？",
                "確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1);

            if (confirm != DialogResult.Yes)
            {
                return;
            }

            Keyword = txtKeywordForMethodSignature.Text;
            IsRegexEnabled = chkEnabledRegExForMethodSignature.Checked;
            IsNotEnabled = chkNot.Checked;
            SaveSettings();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnCloseForm_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void LoadSettings()
        {
            if (!File.Exists(_settingsPath))
            {
                return;
            }

            try
            {
                var json = File.ReadAllText(_settingsPath, Encoding.UTF8);
                var serializer = new JavaScriptSerializer();
                var settings = serializer.Deserialize<FormSettings>(json);
                if (settings == null)
                {
                    return;
                }

                txtKeywordForMethodSignature.Text = settings.Keyword ?? string.Empty;
                chkEnabledRegExForMethodSignature.Checked = settings.IsRegexEnabled;
                chkNot.Checked = settings.IsNotEnabled;
            }
            catch (Exception)
            {
                // 設定読み込み失敗時は既定値のまま
            }
        }

        private void SaveSettings()
        {
            try
            {
                var settings = new FormSettings
                {
                    Keyword = txtKeywordForMethodSignature.Text,
                    IsRegexEnabled = chkEnabledRegExForMethodSignature.Checked,
                    IsNotEnabled = chkNot.Checked
                };
                var serializer = new JavaScriptSerializer();
                var json = serializer.Serialize(settings);
                File.WriteAllText(_settingsPath, json, Encoding.UTF8);
            }
            catch (Exception)
            {
                // 設定保存失敗時は無視
            }
        }

        private static string RemoveMethodsAndImports(string content, string keyword, bool useRegex, bool negateMatch, ILanguageProcessor processor)
        {
            if (string.IsNullOrEmpty(content))
            {
                return content;
            }

            string lineEnding = content.Contains("\r\n") ? "\r\n" : content.Contains("\r") ? "\r" : "\n";
            var lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var methodSpans = processor.FindMethodSpans(lines);
            var linesToRemove = new HashSet<int>();

            for (int i = 0; i < lines.Length; i++)
            {
                if (processor.IsImportLine(lines[i]))
                {
                    linesToRemove.Add(i);
                }
            }

            foreach (var span in methodSpans)
            {
                bool found = useRegex
                    ? Regex.IsMatch(span.RawSignature, keyword)
                    : span.RawSignature.IndexOf(keyword, StringComparison.Ordinal) >= 0;

                bool shouldDelete = negateMatch ? !found : found;

                if (shouldDelete)
                {
                    for (int i = span.StartLine; i <= span.EndLine; i++)
                    {
                        linesToRemove.Add(i);
                    }

                    for (int i = span.EndLine + 1; i < lines.Length; i++)
                    {
                        if (string.IsNullOrWhiteSpace(lines[i]))
                        {
                            linesToRemove.Add(i);
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }

            if (linesToRemove.Count == 0)
            {
                return content;
            }

            var result = new List<string>();
            for (int i = 0; i < lines.Length; i++)
            {
                if (!linesToRemove.Contains(i))
                {
                    result.Add(lines[i]);
                }
            }

            var collapsed = new List<string>();
            bool prevBlank = false;
            for (int i = 0; i < result.Count; i++)
            {
                bool isBlank = string.IsNullOrWhiteSpace(result[i]);
                if (isBlank && prevBlank)
                {
                    continue;
                }

                collapsed.Add(result[i]);
                prevBlank = isBlank;
            }

            return string.Join(lineEnding, collapsed);
        }

        private class FormSettings
        {
            public string Keyword { get; set; } = string.Empty;
            public bool IsRegexEnabled { get; set; }
            public bool IsNotEnabled { get; set; }
        }
    }
}
