namespace SimpleExcelGrep.Forms
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
        protected override void Dispose ( bool disposing )
        {
            if ( disposing && ( components != null ) )
            {
                components.Dispose();
            }
            base.Dispose( disposing );
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent ()
        {
            this.components = new System.ComponentModel.Container();
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
            this.colFilePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSheetName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCellPosition = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cmbKeyword = new System.Windows.Forms.ComboBox();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.lblFolder = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblKeyword = new System.Windows.Forms.Label();
            this.cmbFolderPath = new System.Windows.Forms.ComboBox();
            this.colCellValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.grdResults = new System.Windows.Forms.DataGridView();
            this.lblIgnoreFileSize = new System.Windows.Forms.Label();
            this.txtIgnoreFileSizeMB = new System.Windows.Forms.TextBox();
            this.lblIgnoreFileSizeUnit = new System.Windows.Forms.Label();
            this.btnLoadTsv = new System.Windows.Forms.Button();
            this.txtCellAddress = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.chkCellMode = new System.Windows.Forms.CheckBox();
            this.chkSearchSubDir = new System.Windows.Forms.CheckBox();
            this.chkEnableLog = new System.Windows.Forms.CheckBox();
            this.chkEnableInvisibleSheet = new System.Windows.Forms.CheckBox();
            this.chkDblClickToOpen = new System.Windows.Forms.CheckBox();
            this.chkCollectStrInShape = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.nudParallelism)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdResults)).BeginInit();
            this.SuspendLayout();
            // 
            // btnStartSearch
            // 
            this.btnStartSearch.Location = new System.Drawing.Point(11, 207);
            this.btnStartSearch.Name = "btnStartSearch";
            this.btnStartSearch.Size = new System.Drawing.Size(94, 30);
            this.btnStartSearch.TabIndex = 10;
            this.btnStartSearch.Text = "検索開始";
            this.btnStartSearch.UseVisualStyleBackColor = true;
            // 
            // btnCancelSearch
            // 
            this.btnCancelSearch.Enabled = false;
            this.btnCancelSearch.Location = new System.Drawing.Point(117, 208);
            this.btnCancelSearch.Name = "btnCancelSearch";
            this.btnCancelSearch.Size = new System.Drawing.Size(100, 28);
            this.btnCancelSearch.TabIndex = 11;
            this.btnCancelSearch.Text = "検索中止";
            this.btnCancelSearch.UseVisualStyleBackColor = true;
            // 
            // chkRealTimeDisplay
            // 
            this.chkRealTimeDisplay.AutoSize = true;
            this.chkRealTimeDisplay.Checked = true;
            this.chkRealTimeDisplay.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkRealTimeDisplay.Location = new System.Drawing.Point(233, 213);
            this.chkRealTimeDisplay.Name = "chkRealTimeDisplay";
            this.chkRealTimeDisplay.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.chkRealTimeDisplay.Size = new System.Drawing.Size(106, 16);
            this.chkRealTimeDisplay.TabIndex = 12;
            this.chkRealTimeDisplay.Text = "リアルタイム表示";
            this.chkRealTimeDisplay.UseVisualStyleBackColor = true;
            // 
            // lblParallelism
            // 
            this.lblParallelism.AutoSize = true;
            this.lblParallelism.Location = new System.Drawing.Point(350, 214);
            this.lblParallelism.Name = "lblParallelism";
            this.lblParallelism.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.lblParallelism.Size = new System.Drawing.Size(48, 12);
            this.lblParallelism.TabIndex = 13;
            this.lblParallelism.Text = "並列数:";
            // 
            // nudParallelism
            // 
            this.nudParallelism.Location = new System.Drawing.Point(402, 212);
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
            this.nudParallelism.TabIndex = 14;
            this.nudParallelism.Value = new decimal(new int[] {
            4,
            0,
            0,
            0});
            // 
            // chkSearchShapes
            // 
            this.chkSearchShapes.AutoSize = true;
            this.chkSearchShapes.Location = new System.Drawing.Point(256, 178);
            this.chkSearchShapes.Name = "chkSearchShapes";
            this.chkSearchShapes.Size = new System.Drawing.Size(81, 16);
            this.chkSearchShapes.TabIndex = 9;
            this.chkSearchShapes.Text = "図形も検索";
            this.chkSearchShapes.UseVisualStyleBackColor = true;
            // 
            // chkFirstHitOnly
            // 
            this.chkFirstHitOnly.AutoSize = true;
            this.chkFirstHitOnly.Location = new System.Drawing.Point(133, 178);
            this.chkFirstHitOnly.Name = "chkFirstHitOnly";
            this.chkFirstHitOnly.Size = new System.Drawing.Size(102, 16);
            this.chkFirstHitOnly.TabIndex = 8;
            this.chkFirstHitOnly.Text = "最初のヒットのみ";
            this.chkFirstHitOnly.UseVisualStyleBackColor = true;
            // 
            // chkRegex
            // 
            this.chkRegex.AutoSize = true;
            this.chkRegex.Location = new System.Drawing.Point(13, 178);
            this.chkRegex.Name = "chkRegex";
            this.chkRegex.Size = new System.Drawing.Size(105, 16);
            this.chkRegex.TabIndex = 7;
            this.chkRegex.Text = "正規表現を使用";
            this.chkRegex.UseVisualStyleBackColor = true;
            // 
            // lblIgnoreHint
            // 
            this.lblIgnoreHint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblIgnoreHint.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblIgnoreHint.Location = new System.Drawing.Point(741, 95);
            this.lblIgnoreHint.Name = "lblIgnoreHint";
            this.lblIgnoreHint.Size = new System.Drawing.Size(78, 28);
            this.lblIgnoreHint.TabIndex = 5;
            this.lblIgnoreHint.Text = "カンマ区切り";
            this.lblIgnoreHint.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmbIgnoreKeywords
            // 
            this.cmbIgnoreKeywords.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbIgnoreKeywords.FormattingEnabled = true;
            this.cmbIgnoreKeywords.Location = new System.Drawing.Point(107, 100);
            this.cmbIgnoreKeywords.Name = "cmbIgnoreKeywords";
            this.cmbIgnoreKeywords.Size = new System.Drawing.Size(610, 20);
            this.cmbIgnoreKeywords.TabIndex = 4;
            // 
            // lblIgnore
            // 
            this.lblIgnore.Location = new System.Drawing.Point(11, 95);
            this.lblIgnore.Name = "lblIgnore";
            this.lblIgnore.Size = new System.Drawing.Size(88, 28);
            this.lblIgnore.TabIndex = 61;
            this.lblIgnore.Text = "無視キーワード:";
            this.lblIgnore.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
            // cmbKeyword
            // 
            this.cmbKeyword.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbKeyword.FormattingEnabled = true;
            this.cmbKeyword.Location = new System.Drawing.Point(107, 57);
            this.cmbKeyword.Name = "cmbKeyword";
            this.cmbKeyword.Size = new System.Drawing.Size(606, 20);
            this.cmbKeyword.TabIndex = 2;
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectFolder.Location = new System.Drawing.Point(728, 11);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(57, 22);
            this.btnSelectFolder.TabIndex = 1;
            this.btnSelectFolder.Text = "選択...";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            // 
            // lblFolder
            // 
            this.lblFolder.Location = new System.Drawing.Point(10, 9);
            this.lblFolder.Name = "lblFolder";
            this.lblFolder.Size = new System.Drawing.Size(81, 28);
            this.lblFolder.TabIndex = 60;
            this.lblFolder.Text = "フォルダパス:";
            this.lblFolder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblStatus
            // 
            this.lblStatus.Location = new System.Drawing.Point(10, 242);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(442, 28);
            this.lblStatus.TabIndex = 16;
            this.lblStatus.Text = "準備完了";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblKeyword
            // 
            this.lblKeyword.Location = new System.Drawing.Point(9, 52);
            this.lblKeyword.Name = "lblKeyword";
            this.lblKeyword.Size = new System.Drawing.Size(90, 28);
            this.lblKeyword.TabIndex = 59;
            this.lblKeyword.Text = "検索キーワード:";
            this.lblKeyword.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmbFolderPath
            // 
            this.cmbFolderPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFolderPath.FormattingEnabled = true;
            this.cmbFolderPath.Location = new System.Drawing.Point(107, 13);
            this.cmbFolderPath.Name = "cmbFolderPath";
            this.cmbFolderPath.Size = new System.Drawing.Size(610, 20);
            this.cmbFolderPath.TabIndex = 0;
            // 
            // colCellValue
            // 
            this.colCellValue.HeaderText = "セルの値";
            this.colCellValue.Name = "colCellValue";
            this.colCellValue.ReadOnly = true;
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
            this.grdResults.Location = new System.Drawing.Point(11, 278);
            this.grdResults.Name = "grdResults";
            this.grdResults.ReadOnly = true;
            this.grdResults.RowTemplate.Height = 21;
            this.grdResults.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grdResults.Size = new System.Drawing.Size(824, 202);
            this.grdResults.TabIndex = 17;
            // 
            // lblIgnoreFileSize
            // 
            this.lblIgnoreFileSize.Location = new System.Drawing.Point(11, 137);
            this.lblIgnoreFileSize.Name = "lblIgnoreFileSize";
            this.lblIgnoreFileSize.Size = new System.Drawing.Size(116, 28);
            this.lblIgnoreFileSize.TabIndex = 62;
            this.lblIgnoreFileSize.Text = "無視ファイルサイズ:";
            this.lblIgnoreFileSize.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtIgnoreFileSizeMB
            // 
            this.txtIgnoreFileSizeMB.Location = new System.Drawing.Point(133, 142);
            this.txtIgnoreFileSizeMB.Name = "txtIgnoreFileSizeMB";
            this.txtIgnoreFileSizeMB.Size = new System.Drawing.Size(60, 19);
            this.txtIgnoreFileSizeMB.TabIndex = 5;
            this.txtIgnoreFileSizeMB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // lblIgnoreFileSizeUnit
            // 
            this.lblIgnoreFileSizeUnit.AutoSize = true;
            this.lblIgnoreFileSizeUnit.Location = new System.Drawing.Point(199, 145);
            this.lblIgnoreFileSizeUnit.Name = "lblIgnoreFileSizeUnit";
            this.lblIgnoreFileSizeUnit.Size = new System.Drawing.Size(30, 12);
            this.lblIgnoreFileSizeUnit.TabIndex = 6;
            this.lblIgnoreFileSizeUnit.Text = "(MB)";
            // 
            // btnLoadTsv
            // 
            this.btnLoadTsv.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnLoadTsv.Location = new System.Drawing.Point(728, 208);
            this.btnLoadTsv.Name = "btnLoadTsv";
            this.btnLoadTsv.Size = new System.Drawing.Size(107, 28);
            this.btnLoadTsv.TabIndex = 15;
            this.btnLoadTsv.Text = "TSV読み込み";
            this.btnLoadTsv.UseVisualStyleBackColor = true;
            // 
            // txtCellAddress
            // 
            this.txtCellAddress.Enabled = false;
            this.txtCellAddress.Location = new System.Drawing.Point(338, 142);
            this.txtCellAddress.Name = "txtCellAddress";
            this.txtCellAddress.Size = new System.Drawing.Size(180, 19);
            this.txtCellAddress.TabIndex = 63;
            this.toolTip1.SetToolTip(this.txtCellAddress, "A1形式で指定。半角カンマで複数指定可。");
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(254, 137);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(85, 28);
            this.label2.TabIndex = 65;
            this.label2.Text = "指定セルモード:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // chkCellMode
            // 
            this.chkCellMode.AutoSize = true;
            this.chkCellMode.Location = new System.Drawing.Point(524, 145);
            this.chkCellMode.Name = "chkCellMode";
            this.chkCellMode.Size = new System.Drawing.Size(68, 16);
            this.chkCellMode.TabIndex = 66;
            this.chkCellMode.Text = "ON/OFF";
            this.chkCellMode.UseVisualStyleBackColor = true;
            this.chkCellMode.CheckedChanged += new System.EventHandler(this.ChkCellMode_CheckedChanged);
            // 
            // chkSearchSubDir
            // 
            this.chkSearchSubDir.AutoSize = true;
            this.chkSearchSubDir.Location = new System.Drawing.Point(107, 35);
            this.chkSearchSubDir.Name = "chkSearchSubDir";
            this.chkSearchSubDir.Size = new System.Drawing.Size(111, 16);
            this.chkSearchSubDir.TabIndex = 67;
            this.chkSearchSubDir.Text = "サブフォルダも対象";
            this.chkSearchSubDir.UseVisualStyleBackColor = true;
            // 
            // chkEnableLog
            // 
            this.chkEnableLog.AutoSize = true;
            this.chkEnableLog.Location = new System.Drawing.Point(352, 178);
            this.chkEnableLog.Name = "chkEnableLog";
            this.chkEnableLog.Size = new System.Drawing.Size(75, 16);
            this.chkEnableLog.TabIndex = 68;
            this.chkEnableLog.Text = "ログを出力";
            this.chkEnableLog.UseVisualStyleBackColor = true;
            // 
            // chkEnableInvisibleSheet
            // 
            this.chkEnableInvisibleSheet.AutoSize = true;
            this.chkEnableInvisibleSheet.Location = new System.Drawing.Point(443, 178);
            this.chkEnableInvisibleSheet.Name = "chkEnableInvisibleSheet";
            this.chkEnableInvisibleSheet.Size = new System.Drawing.Size(121, 16);
            this.chkEnableInvisibleSheet.TabIndex = 69;
            this.chkEnableInvisibleSheet.Text = "非表示シートも対象";
            this.chkEnableInvisibleSheet.UseVisualStyleBackColor = true;
            // 
            // chkDblClickToOpen
            // 
            this.chkDblClickToOpen.AutoSize = true;
            this.chkDblClickToOpen.Location = new System.Drawing.Point(256, 35);
            this.chkDblClickToOpen.Name = "chkDblClickToOpen";
            this.chkDblClickToOpen.Size = new System.Drawing.Size(153, 16);
            this.chkDblClickToOpen.TabIndex = 70;
            this.chkDblClickToOpen.Text = "ダブルクリックでファイルを開く";
            this.chkDblClickToOpen.UseVisualStyleBackColor = true;
            // 
            // chkCollectStrInShape
            // 
            this.chkCollectStrInShape.AutoSize = true;
            this.chkCollectStrInShape.Location = new System.Drawing.Point(633, 144);
            this.chkCollectStrInShape.Name = "chkCollectStrInShape";
            this.chkCollectStrInShape.Size = new System.Drawing.Size(148, 16);
            this.chkCollectStrInShape.TabIndex = 71;
            this.chkCollectStrInShape.Text = "図形内文字列収集モード";
            this.chkCollectStrInShape.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(844, 489);
            this.Controls.Add(this.chkCollectStrInShape);
            this.Controls.Add(this.chkDblClickToOpen);
            this.Controls.Add(this.chkEnableInvisibleSheet);
            this.Controls.Add(this.chkEnableLog);
            this.Controls.Add(this.chkSearchSubDir);
            this.Controls.Add(this.chkCellMode);
            this.Controls.Add(this.txtCellAddress);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnLoadTsv);
            this.Controls.Add(this.lblIgnoreFileSizeUnit);
            this.Controls.Add(this.txtIgnoreFileSizeMB);
            this.Controls.Add(this.lblIgnoreFileSize);
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
            this.Name = "MainForm";
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
        private System.Windows.Forms.DataGridViewTextBoxColumn colFilePath;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFileName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSheetName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCellPosition;
        private System.Windows.Forms.ComboBox cmbKeyword;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.Label lblFolder;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblKeyword;
        private System.Windows.Forms.ComboBox cmbFolderPath;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCellValue;
        private System.Windows.Forms.DataGridView grdResults;
        private System.Windows.Forms.Label lblIgnoreFileSize;
        private System.Windows.Forms.TextBox txtIgnoreFileSizeMB;
        private System.Windows.Forms.Label lblIgnoreFileSizeUnit;
        private System.Windows.Forms.Button btnLoadTsv;
        private System.Windows.Forms.TextBox txtCellAddress;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox chkCellMode;
        private System.Windows.Forms.CheckBox chkSearchSubDir;
        private System.Windows.Forms.CheckBox chkEnableLog;
        private System.Windows.Forms.CheckBox chkEnableInvisibleSheet;
        private System.Windows.Forms.CheckBox chkDblClickToOpen;
        private System.Windows.Forms.CheckBox chkCollectStrInShape;
    }
}