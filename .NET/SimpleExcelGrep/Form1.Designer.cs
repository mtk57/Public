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
            this.btnStartSearch = new System.Windows.Forms.Button();
            this.btnCancelSearch = new System.Windows.Forms.Button();
            this.chkRealTimeDisplay = new System.Windows.Forms.CheckBox();
            this.lblParallelism = new System.Windows.Forms.Label();
            this.nudParallelism = new System.Windows.Forms.NumericUpDown();
            this.chkSearchShapes = new System.Windows.Forms.CheckBox();
            this.chkFirstHitOnly = new System.Windows.Forms.CheckBox();
            this.chkRegex = new System.Windows.Forms.CheckBox();
            this.lblIgnoreHint = new System.Windows.Forms.Label();
            this.cmbIgnoreKeywords = new System.Windows.Forms.ComboBox();
            this.lblIgnore = new System.Windows.Forms.Label();
            this.cmbKeyword = new System.Windows.Forms.ComboBox();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.lblFolder = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.colCellValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCellPosition = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSheetName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFilePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblKeyword = new System.Windows.Forms.Label();
            this.cmbFolderPath = new System.Windows.Forms.ComboBox();
            this.grdResults = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.nudParallelism)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdResults)).BeginInit();
            this.SuspendLayout();
            // 
            // btnStartSearch
            // 
            this.btnStartSearch.Location = new System.Drawing.Point(10, 178);
            this.btnStartSearch.Name = "btnStartSearch";
            this.btnStartSearch.Size = new System.Drawing.Size(94, 30);
            this.btnStartSearch.TabIndex = 43;
            this.btnStartSearch.Text = "検索開始";
            this.btnStartSearch.UseVisualStyleBackColor = true;
            // 
            // btnCancelSearch
            // 
            this.btnCancelSearch.Enabled = false;
            this.btnCancelSearch.Location = new System.Drawing.Point(116, 179);
            this.btnCancelSearch.Name = "btnCancelSearch";
            this.btnCancelSearch.Size = new System.Drawing.Size(100, 28);
            this.btnCancelSearch.TabIndex = 44;
            this.btnCancelSearch.Text = "検索中止";
            this.btnCancelSearch.UseVisualStyleBackColor = true;
            // 
            // chkRealTimeDisplay
            // 
            this.chkRealTimeDisplay.AutoSize = true;
            this.chkRealTimeDisplay.Checked = true;
            this.chkRealTimeDisplay.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkRealTimeDisplay.Location = new System.Drawing.Point(222, 179);
            this.chkRealTimeDisplay.Name = "chkRealTimeDisplay";
            this.chkRealTimeDisplay.Padding = new System.Windows.Forms.Padding(10, 5, 0, 0);
            this.chkRealTimeDisplay.Size = new System.Drawing.Size(111, 21);
            this.chkRealTimeDisplay.TabIndex = 45;
            this.chkRealTimeDisplay.Text = "リアルタイム表示";
            this.chkRealTimeDisplay.UseVisualStyleBackColor = true;
            // 
            // lblParallelism
            // 
            this.lblParallelism.AutoSize = true;
            this.lblParallelism.Location = new System.Drawing.Point(339, 179);
            this.lblParallelism.Name = "lblParallelism";
            this.lblParallelism.Padding = new System.Windows.Forms.Padding(10, 5, 0, 0);
            this.lblParallelism.Size = new System.Drawing.Size(53, 17);
            this.lblParallelism.TabIndex = 46;
            this.lblParallelism.Text = "並列数:";
            // 
            // nudParallelism
            // 
            this.nudParallelism.Location = new System.Drawing.Point(398, 180);
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
            this.nudParallelism.TabIndex = 47;
            this.nudParallelism.Value = new decimal(new int[] {
            4,
            0,
            0,
            0});
            // 
            // chkSearchShapes
            // 
            this.chkSearchShapes.AutoSize = true;
            this.chkSearchShapes.Location = new System.Drawing.Point(255, 149);
            this.chkSearchShapes.Name = "chkSearchShapes";
            this.chkSearchShapes.Size = new System.Drawing.Size(81, 16);
            this.chkSearchShapes.TabIndex = 42;
            this.chkSearchShapes.Text = "図形も検索";
            this.chkSearchShapes.UseVisualStyleBackColor = true;
            // 
            // chkFirstHitOnly
            // 
            this.chkFirstHitOnly.AutoSize = true;
            this.chkFirstHitOnly.Location = new System.Drawing.Point(132, 149);
            this.chkFirstHitOnly.Name = "chkFirstHitOnly";
            this.chkFirstHitOnly.Size = new System.Drawing.Size(102, 16);
            this.chkFirstHitOnly.TabIndex = 41;
            this.chkFirstHitOnly.Text = "最初のヒットのみ";
            this.chkFirstHitOnly.UseVisualStyleBackColor = true;
            // 
            // chkRegex
            // 
            this.chkRegex.AutoSize = true;
            this.chkRegex.Location = new System.Drawing.Point(12, 149);
            this.chkRegex.Name = "chkRegex";
            this.chkRegex.Size = new System.Drawing.Size(105, 16);
            this.chkRegex.TabIndex = 40;
            this.chkRegex.Text = "正規表現を使用";
            this.chkRegex.UseVisualStyleBackColor = true;
            // 
            // lblIgnoreHint
            // 
            this.lblIgnoreHint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblIgnoreHint.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblIgnoreHint.Location = new System.Drawing.Point(740, 94);
            this.lblIgnoreHint.Name = "lblIgnoreHint";
            this.lblIgnoreHint.Size = new System.Drawing.Size(78, 28);
            this.lblIgnoreHint.TabIndex = 39;
            this.lblIgnoreHint.Text = "カンマ区切り";
            this.lblIgnoreHint.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmbIgnoreKeywords
            // 
            this.cmbIgnoreKeywords.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbIgnoreKeywords.FormattingEnabled = true;
            this.cmbIgnoreKeywords.Location = new System.Drawing.Point(106, 99);
            this.cmbIgnoreKeywords.Name = "cmbIgnoreKeywords";
            this.cmbIgnoreKeywords.Size = new System.Drawing.Size(610, 20);
            this.cmbIgnoreKeywords.TabIndex = 38;
            // 
            // lblIgnore
            // 
            this.lblIgnore.Location = new System.Drawing.Point(10, 94);
            this.lblIgnore.Name = "lblIgnore";
            this.lblIgnore.Size = new System.Drawing.Size(88, 28);
            this.lblIgnore.TabIndex = 37;
            this.lblIgnore.Text = "無視キーワード:";
            this.lblIgnore.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmbKeyword
            // 
            this.cmbKeyword.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbKeyword.FormattingEnabled = true;
            this.cmbKeyword.Location = new System.Drawing.Point(106, 56);
            this.cmbKeyword.Name = "cmbKeyword";
            this.cmbKeyword.Size = new System.Drawing.Size(606, 20);
            this.cmbKeyword.TabIndex = 36;
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFolder.Location = new System.Drawing.Point(727, 10);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(57, 22);
            this.btnSelectFolder.TabIndex = 34;
            this.btnSelectFolder.Text = "選択...";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            // 
            // lblFolder
            // 
            this.lblFolder.Location = new System.Drawing.Point(9, 8);
            this.lblFolder.Name = "lblFolder";
            this.lblFolder.Size = new System.Drawing.Size(81, 28);
            this.lblFolder.TabIndex = 33;
            this.lblFolder.Text = "フォルダパス:";
            this.lblFolder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblStatus
            // 
            this.lblStatus.Location = new System.Drawing.Point(8, 216);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(376, 28);
            this.lblStatus.TabIndex = 31;
            this.lblStatus.Text = "準備完了";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // colCellValue
            // 
            this.colCellValue.HeaderText = "セルの値";
            this.colCellValue.Name = "colCellValue";
            this.colCellValue.ReadOnly = true;
            // 
            // colCellPosition
            // 
            this.colCellPosition.HeaderText = "セル位置";
            this.colCellPosition.Name = "colCellPosition";
            this.colCellPosition.ReadOnly = true;
            // 
            // colSheetName
            // 
            this.colSheetName.HeaderText = "シート名";
            this.colSheetName.Name = "colSheetName";
            this.colSheetName.ReadOnly = true;
            // 
            // colFileName
            // 
            this.colFileName.HeaderText = "ファイル名";
            this.colFileName.Name = "colFileName";
            this.colFileName.ReadOnly = true;
            // 
            // colFilePath
            // 
            this.colFilePath.HeaderText = "ファイルパス";
            this.colFilePath.Name = "colFilePath";
            this.colFilePath.ReadOnly = true;
            // 
            // lblKeyword
            // 
            this.lblKeyword.Location = new System.Drawing.Point(8, 51);
            this.lblKeyword.Name = "lblKeyword";
            this.lblKeyword.Size = new System.Drawing.Size(90, 28);
            this.lblKeyword.TabIndex = 35;
            this.lblKeyword.Text = "検索キーワード:";
            this.lblKeyword.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmbFolderPath
            // 
            this.cmbFolderPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFolderPath.FormattingEnabled = true;
            this.cmbFolderPath.Location = new System.Drawing.Point(106, 12);
            this.cmbFolderPath.Name = "cmbFolderPath";
            this.cmbFolderPath.Size = new System.Drawing.Size(610, 20);
            this.cmbFolderPath.TabIndex = 32;
            // 
            // grdResults
            // 
            this.grdResults.AllowUserToAddRows = false;
            this.grdResults.AllowUserToDeleteRows = false;
            this.grdResults.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grdResults.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.grdResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdResults.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colFilePath,
            this.colFileName,
            this.colSheetName,
            this.colCellPosition,
            this.colCellValue});
            this.grdResults.Location = new System.Drawing.Point(10, 250);
            this.grdResults.Name = "grdResults";
            this.grdResults.ReadOnly = true;
            this.grdResults.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grdResults.Size = new System.Drawing.Size(824, 229);
            this.grdResults.TabIndex = 30;

            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(844, 489);
            this.Controls.Add(this.btnStartSearch);
            this.Controls.Add(this.btnCancelSearch);
            this.Controls.Add(this.chkRealTimeDisplay);
            this.Controls.Add(this.lblParallelism);
            this.Controls.Add(this.nudParallelism);
            this.Controls.Add(this.chkSearchShapes);
            this.Controls.Add(this.chkFirstHitOnly);
            this.Controls.Add(this.chkRegex);
            this.Controls.Add(this.lblIgnoreHint);
            this.Controls.Add(this.cmbIgnoreKeywords);
            this.Controls.Add(this.lblIgnore);
            this.Controls.Add(this.cmbKeyword);
            this.Controls.Add(this.btnSelectFolder);
            this.Controls.Add(this.lblFolder);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.lblKeyword);
            this.Controls.Add(this.cmbFolderPath);
            this.Controls.Add(this.grdResults);
            this.MinimumSize = new System.Drawing.Size(640, 446);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Simple Excel Grep";
            ((System.ComponentModel.ISupportInitialize)(this.nudParallelism)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdResults)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStartSearch;
        private System.Windows.Forms.Button btnCancelSearch;
        private System.Windows.Forms.CheckBox chkRealTimeDisplay;
        private System.Windows.Forms.Label lblParallelism;
        private System.Windows.Forms.NumericUpDown nudParallelism;
        private System.Windows.Forms.CheckBox chkSearchShapes;
        private System.Windows.Forms.CheckBox chkFirstHitOnly;
        private System.Windows.Forms.CheckBox chkRegex;
        private System.Windows.Forms.Label lblIgnoreHint;
        private System.Windows.Forms.ComboBox cmbIgnoreKeywords;
        private System.Windows.Forms.Label lblIgnore;
        private System.Windows.Forms.ComboBox cmbKeyword;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.Label lblFolder;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCellValue;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCellPosition;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSheetName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFileName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFilePath;
        private System.Windows.Forms.Label lblKeyword;
        private System.Windows.Forms.ComboBox cmbFolderPath;
        private System.Windows.Forms.DataGridView grdResults;
    }
}