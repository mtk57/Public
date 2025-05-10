namespace SimpleExcelGrep
{
     partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.mainLayout = new System.Windows.Forms.TableLayoutPanel();
            this.lblFolder = new System.Windows.Forms.Label();
            this.cmbFolderPath = new System.Windows.Forms.ComboBox();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.lblKeyword = new System.Windows.Forms.Label();
            this.cmbKeyword = new System.Windows.Forms.ComboBox();
            this.chkRegex = new System.Windows.Forms.CheckBox();
            this.chkFirstHitOnly = new System.Windows.Forms.CheckBox();
            this.lblIgnore = new System.Windows.Forms.Label();
            this.cmbIgnoreKeywords = new System.Windows.Forms.ComboBox();
            this.lblIgnoreHint = new System.Windows.Forms.Label();
            this.buttonPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnStartSearch = new System.Windows.Forms.Button();
            this.btnCancelSearch = new System.Windows.Forms.Button();
            this.chkRealTimeDisplay = new System.Windows.Forms.CheckBox();
            this.lblParallelism = new System.Windows.Forms.Label();
            this.nudParallelism = new System.Windows.Forms.NumericUpDown();
            this.lblStatus = new System.Windows.Forms.Label();
            this.grdResults = new System.Windows.Forms.DataGridView();
            this.colFilePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSheetName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCellPosition = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCellValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mainLayout.SuspendLayout();
            this.buttonPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudParallelism)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdResults)).BeginInit();
            this.SuspendLayout();
            // 
            // mainLayout
            // 
            this.mainLayout.ColumnCount = 3;
            this.mainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.mainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70F));
            this.mainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.mainLayout.Controls.Add(this.lblFolder, 0, 0);
            this.mainLayout.Controls.Add(this.cmbFolderPath, 1, 0);
            this.mainLayout.Controls.Add(this.btnSelectFolder, 2, 0);
            this.mainLayout.Controls.Add(this.lblKeyword, 0, 1);
            this.mainLayout.Controls.Add(this.cmbKeyword, 1, 1);
            this.mainLayout.Controls.Add(this.chkRegex, 1, 2);
            this.mainLayout.Controls.Add(this.chkFirstHitOnly, 1, 2);
            this.mainLayout.Controls.Add(this.lblIgnore, 0, 3);
            this.mainLayout.Controls.Add(this.cmbIgnoreKeywords, 1, 3);
            this.mainLayout.Controls.Add(this.lblIgnoreHint, 2, 3);
            this.mainLayout.Controls.Add(this.buttonPanel, 1, 4);
            this.mainLayout.Controls.Add(this.lblStatus, 1, 5);
            this.mainLayout.Controls.Add(this.grdResults, 0, 6);
            this.mainLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainLayout.Location = new System.Drawing.Point(0, 0);
            this.mainLayout.Name = "mainLayout";
            this.mainLayout.Padding = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.mainLayout.RowCount = 7;
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 37F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.mainLayout.Size = new System.Drawing.Size(784, 518);
            this.mainLayout.TabIndex = 0;
            // 
            // lblFolder
            // 
            this.lblFolder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFolder.Location = new System.Drawing.Point(13, 9);
            this.lblFolder.Name = "lblFolder";
            this.lblFolder.Size = new System.Drawing.Size(108, 28);
            this.lblFolder.TabIndex = 0;
            this.lblFolder.Text = "フォルダパス:";
            this.lblFolder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmbFolderPath
            // 
            this.cmbFolderPath.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cmbFolderPath.FormattingEnabled = true;
            this.cmbFolderPath.Location = new System.Drawing.Point(127, 12);
            this.cmbFolderPath.Name = "cmbFolderPath";
            this.cmbFolderPath.Size = new System.Drawing.Size(528, 20);
            this.cmbFolderPath.TabIndex = 1;
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnSelectFolder.Location = new System.Drawing.Point(661, 12);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(110, 22);
            this.btnSelectFolder.TabIndex = 2;
            this.btnSelectFolder.Text = "選択...";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            // 
            // lblKeyword
            // 
            this.lblKeyword.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblKeyword.Location = new System.Drawing.Point(13, 37);
            this.lblKeyword.Name = "lblKeyword";
            this.lblKeyword.Size = new System.Drawing.Size(108, 28);
            this.lblKeyword.TabIndex = 3;
            this.lblKeyword.Text = "検索キーワード:";
            this.lblKeyword.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmbKeyword
            // 
            this.cmbKeyword.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cmbKeyword.FormattingEnabled = true;
            this.cmbKeyword.Location = new System.Drawing.Point(127, 40);
            this.cmbKeyword.Name = "cmbKeyword";
            this.cmbKeyword.Size = new System.Drawing.Size(528, 20);
            this.cmbKeyword.TabIndex = 4;
            this.cmbKeyword.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbKeyword_KeyDown);
            // 
            // chkRegex
            // 
            this.chkRegex.AutoSize = true;
            this.chkRegex.Dock = System.Windows.Forms.DockStyle.Left;
            this.chkRegex.Location = new System.Drawing.Point(661, 68);
            this.chkRegex.Name = "chkRegex";
            this.chkRegex.Size = new System.Drawing.Size(105, 22);
            this.chkRegex.TabIndex = 5;
            this.chkRegex.Text = "正規表現を使用";
            this.chkRegex.UseVisualStyleBackColor = true;
            // 
            // chkFirstHitOnly
            // 
            this.chkFirstHitOnly.AutoSize = true;
            this.chkFirstHitOnly.Dock = System.Windows.Forms.DockStyle.Right;
            this.chkFirstHitOnly.Location = new System.Drawing.Point(429, 68);
            this.chkFirstHitOnly.Name = "chkFirstHitOnly";
            this.chkFirstHitOnly.Size = new System.Drawing.Size(226, 22);
            this.chkFirstHitOnly.TabIndex = 12;
            this.chkFirstHitOnly.Text = "最初のヒットのみ（より高速に検索できます）";
            this.chkFirstHitOnly.UseVisualStyleBackColor = true;
            // 
            // lblIgnore
            // 
            this.lblIgnore.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblIgnore.Location = new System.Drawing.Point(13, 93);
            this.lblIgnore.Name = "lblIgnore";
            this.lblIgnore.Size = new System.Drawing.Size(108, 28);
            this.lblIgnore.TabIndex = 6;
            this.lblIgnore.Text = "無視キーワード:";
            this.lblIgnore.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmbIgnoreKeywords
            // 
            this.cmbIgnoreKeywords.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cmbIgnoreKeywords.FormattingEnabled = true;
            this.cmbIgnoreKeywords.Location = new System.Drawing.Point(127, 96);
            this.cmbIgnoreKeywords.Name = "cmbIgnoreKeywords";
            this.cmbIgnoreKeywords.Size = new System.Drawing.Size(528, 20);
            this.cmbIgnoreKeywords.TabIndex = 7;
            // 
            // lblIgnoreHint
            // 
            this.lblIgnoreHint.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblIgnoreHint.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblIgnoreHint.Location = new System.Drawing.Point(661, 93);
            this.lblIgnoreHint.Name = "lblIgnoreHint";
            this.lblIgnoreHint.Size = new System.Drawing.Size(110, 28);
            this.lblIgnoreHint.TabIndex = 8;
            this.lblIgnoreHint.Text = "カンマ区切り";
            this.lblIgnoreHint.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // buttonPanel
            // 
            this.buttonPanel.Controls.Add(this.btnStartSearch);
            this.buttonPanel.Controls.Add(this.btnCancelSearch);
            this.buttonPanel.Controls.Add(this.chkRealTimeDisplay);
            this.buttonPanel.Controls.Add(this.lblParallelism);
            this.buttonPanel.Controls.Add(this.nudParallelism);
            this.buttonPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.buttonPanel.Location = new System.Drawing.Point(127, 124);
            this.buttonPanel.Name = "buttonPanel";
            this.buttonPanel.Size = new System.Drawing.Size(528, 31);
            this.buttonPanel.TabIndex = 9;
            // 
            // btnStartSearch
            // 
            this.btnStartSearch.Location = new System.Drawing.Point(3, 3);
            this.btnStartSearch.Name = "btnStartSearch";
            this.btnStartSearch.Size = new System.Drawing.Size(100, 28);
            this.btnStartSearch.TabIndex = 0;
            this.btnStartSearch.Text = "検索開始";
            this.btnStartSearch.UseVisualStyleBackColor = true;
            // 
            // btnCancelSearch
            // 
            this.btnCancelSearch.Enabled = false;
            this.btnCancelSearch.Location = new System.Drawing.Point(109, 3);
            this.btnCancelSearch.Name = "btnCancelSearch";
            this.btnCancelSearch.Size = new System.Drawing.Size(100, 28);
            this.btnCancelSearch.TabIndex = 1;
            this.btnCancelSearch.Text = "検索中止";
            this.btnCancelSearch.UseVisualStyleBackColor = true;
            // 
            // chkRealTimeDisplay
            // 
            this.chkRealTimeDisplay.AutoSize = true;
            this.chkRealTimeDisplay.Checked = true;
            this.chkRealTimeDisplay.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkRealTimeDisplay.Location = new System.Drawing.Point(215, 3);
            this.chkRealTimeDisplay.Name = "chkRealTimeDisplay";
            this.chkRealTimeDisplay.Padding = new System.Windows.Forms.Padding(10, 5, 0, 0);
            this.chkRealTimeDisplay.Size = new System.Drawing.Size(111, 21);
            this.chkRealTimeDisplay.TabIndex = 2;
            this.chkRealTimeDisplay.Text = "リアルタイム表示";
            this.chkRealTimeDisplay.UseVisualStyleBackColor = true;
            // 
            // lblParallelism
            // 
            this.lblParallelism.AutoSize = true;
            this.lblParallelism.Location = new System.Drawing.Point(332, 0);
            this.lblParallelism.Name = "lblParallelism";
            this.lblParallelism.Padding = new System.Windows.Forms.Padding(10, 5, 0, 0);
            this.lblParallelism.Size = new System.Drawing.Size(53, 17);
            this.lblParallelism.TabIndex = 3;
            this.lblParallelism.Text = "並列数:";
            // 
            // nudParallelism
            // 
            this.nudParallelism.Location = new System.Drawing.Point(391, 3);
            this.nudParallelism.Maximum = new decimal(new int[] {
            32,
            0,
            0,
            0});
            this.nudParallelism.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudParallelism.Name = "nudParallelism";
            this.nudParallelism.Size = new System.Drawing.Size(50, 19);
            this.nudParallelism.TabIndex = 4;
            this.nudParallelism.Value = new decimal(new int[] {
            4,
            0,
            0,
            0});
            // 
            // lblStatus
            // 
            this.lblStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblStatus.Location = new System.Drawing.Point(127, 158);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(528, 28);
            this.lblStatus.TabIndex = 10;
            this.lblStatus.Text = "準備完了";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // grdResults
            // 
            this.grdResults.AllowUserToAddRows = false;
            this.grdResults.AllowUserToDeleteRows = false;
            this.grdResults.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.grdResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdResults.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colFilePath,
            this.colFileName,
            this.colSheetName,
            this.colCellPosition,
            this.colCellValue});
            this.mainLayout.SetColumnSpan(this.grdResults, 3);
            this.grdResults.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grdResults.Location = new System.Drawing.Point(13, 189);
            this.grdResults.Name = "grdResults";
            this.grdResults.ReadOnly = true;
            this.grdResults.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grdResults.Size = new System.Drawing.Size(758, 317);
            this.grdResults.TabIndex = 11;
            // 
            // colFilePath
            // 
            this.colFilePath.HeaderText = "ファイルパス";
            this.colFilePath.Name = "colFilePath";
            this.colFilePath.ReadOnly = true;
            // 
            // colFileName
            // 
            this.colFileName.HeaderText = "ファイル名";
            this.colFileName.Name = "colFileName";
            this.colFileName.ReadOnly = true;
            // 
            // colSheetName
            // 
            this.colSheetName.HeaderText = "シート名";
            this.colSheetName.Name = "colSheetName";
            this.colSheetName.ReadOnly = true;
            // 
            // colCellPosition
            // 
            this.colCellPosition.HeaderText = "セル位置";
            this.colCellPosition.Name = "colCellPosition";
            this.colCellPosition.ReadOnly = true;
            // 
            // colCellValue
            // 
            this.colCellValue.HeaderText = "セルの値";
            this.colCellValue.Name = "colCellValue";
            this.colCellValue.ReadOnly = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 518);
            this.Controls.Add(this.mainLayout);
            this.MinimumSize = new System.Drawing.Size(640, 446);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Simple Excel Grep";
            this.mainLayout.ResumeLayout(false);
            this.mainLayout.PerformLayout();
            this.buttonPanel.ResumeLayout(false);
            this.buttonPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudParallelism)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdResults)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel mainLayout;
        private System.Windows.Forms.Label lblFolder;
        private System.Windows.Forms.ComboBox cmbFolderPath;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.Label lblKeyword;
        private System.Windows.Forms.ComboBox cmbKeyword;
        private System.Windows.Forms.CheckBox chkRegex;
        private System.Windows.Forms.CheckBox chkFirstHitOnly;
        private System.Windows.Forms.Label lblIgnore;
        private System.Windows.Forms.ComboBox cmbIgnoreKeywords;
        private System.Windows.Forms.Label lblIgnoreHint;
        private System.Windows.Forms.FlowLayoutPanel buttonPanel;
        private System.Windows.Forms.Button btnStartSearch;
        private System.Windows.Forms.Button btnCancelSearch;
        private System.Windows.Forms.CheckBox chkRealTimeDisplay;
        private System.Windows.Forms.Label lblParallelism;
        private System.Windows.Forms.NumericUpDown nudParallelism;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.DataGridView grdResults;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFilePath;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFileName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSheetName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCellPosition;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCellValue;
    }
}