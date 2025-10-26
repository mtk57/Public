using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SimpleMethodCallListCreator
{
    public partial class IgnoreForm : Form
    {
        private readonly List<IgnoreConditionSetting> _conditions = new List<IgnoreConditionSetting>();

        public IgnoreForm ()
        {
            InitializeComponent();
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.MultiSelect = true;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.DefaultValuesNeeded += DataGridView1_DefaultValuesNeeded;
            btnInsertRow.Click += BtnInsertRow_Click;
            btnRemoveRow.Click += BtnRemoveRow_Click;
        }

        public void SetConditions(IEnumerable<IgnoreConditionSetting> conditions)
        {
            dataGridView1.Rows.Clear();
            _conditions.Clear();
            if (conditions == null)
            {
                return;
            }

            foreach (var condition in conditions)
            {
                _conditions.Add(new IgnoreConditionSetting
                {
                    Keyword = condition.Keyword,
                    Rule = condition.Rule,
                    UseRegex = condition.UseRegex,
                    MatchCase = condition.MatchCase
                });
            }

            foreach (var condition in _conditions)
            {
                var index = dataGridView1.Rows.Add();
                var row = dataGridView1.Rows[index];
                row.Cells[clmKeyword.Index].Value = condition.Keyword ?? string.Empty;
                row.Cells[clmMode.Index].Value = RuleToDisplay(condition.Rule);
                row.Cells[clmRegEx.Index].Value = condition.UseRegex;
                row.Cells[clmCase.Index].Value = condition.MatchCase;
            }
        }

        public List<IgnoreConditionSetting> GetConditions()
        {
            var copy = new List<IgnoreConditionSetting>(_conditions.Count);
            foreach (var condition in _conditions)
            {
                copy.Add(new IgnoreConditionSetting
                {
                    Keyword = condition.Keyword,
                    Rule = condition.Rule,
                    UseRegex = condition.UseRegex,
                    MatchCase = condition.MatchCase
                });
            }

            return copy;
        }

        private void BtnInsertRow_Click(object sender, EventArgs e)
        {
            var index = dataGridView1.Rows.Add();
            var row = dataGridView1.Rows[index];
            row.Cells[clmMode.Index].Value = RuleToDisplay(IgnoreRule.StartsWith);
            dataGridView1.CurrentCell = row.Cells[clmKeyword.Index];
            dataGridView1.BeginEdit(true);
        }

        private void DataGridView1_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells[clmMode.Index].Value = RuleToDisplay(IgnoreRule.StartsWith);
            e.Row.Cells[clmRegEx.Index].Value = false;
            e.Row.Cells[clmCase.Index].Value = false;
        }

        private void BtnRemoveRow_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                if (!row.IsNewRow)
                {
                    dataGridView1.Rows.Remove(row);
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing ||
                e.CloseReason == CloseReason.None)
            {
                if (!TryCollectConditions(out var collected))
                {
                    e.Cancel = true;
                    return;
                }

                _conditions.Clear();
                _conditions.AddRange(collected);
                DialogResult = DialogResult.OK;
            }

            base.OnFormClosing(e);
        }

        private bool TryCollectConditions(out List<IgnoreConditionSetting> result)
        {
            result = new List<IgnoreConditionSetting>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                var keyword = Convert.ToString(row.Cells[clmKeyword.Index].Value)?.Trim() ?? string.Empty;
                var modeValue = row.Cells[clmMode.Index].Value;
                var useRegex = Convert.ToBoolean(row.Cells[clmRegEx.Index].Value ?? false);
                var matchCase = Convert.ToBoolean(row.Cells[clmCase.Index].Value ?? false);

                if (keyword.Length == 0)
                {
                    continue;
                }

                var rule = ParseRule(modeValue);
                if (useRegex)
                {
                    try
                    {
                        var options = matchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                        _ = new Regex(keyword, options);
                    }
                    catch (ArgumentException ex)
                    {
                        MessageBox.Show(this, $"正規表現が不正です。\n{ex.Message}", "入力エラー",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        dataGridView1.CurrentCell = row.Cells[clmKeyword.Index];
                        dataGridView1.BeginEdit(true);
                        return false;
                    }
                }

                result.Add(new IgnoreConditionSetting
                {
                    Keyword = keyword,
                    Rule = rule,
                    UseRegex = useRegex,
                    MatchCase = matchCase
                });
            }

            return true;
        }

        private static string RuleToDisplay(IgnoreRule rule)
        {
            switch (rule)
            {
                case IgnoreRule.EndsWith:
                    return "終わる";
                case IgnoreRule.Contains:
                    return "含む";
                case IgnoreRule.StartsWith:
                default:
                    return "始まる";
            }
        }

        private static IgnoreRule ParseRule(object value)
        {
            var text = Convert.ToString(value);
            switch (text)
            {
                case "終わる":
                    return IgnoreRule.EndsWith;
                case "含む":
                    return IgnoreRule.Contains;
                default:
                    return IgnoreRule.StartsWith;
            }
        }
    }
}
