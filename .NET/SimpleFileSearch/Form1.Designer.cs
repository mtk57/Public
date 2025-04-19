namespace SimpleFileSearch
{
    partial class MainForm
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose ( bool disposing )
        {
            if ( disposing && ( components != null ) )
            {
                components.Dispose();
            }
            base.Dispose( disposing );
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent ()
        {
            this.btnBrowse = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbKeyword = new System.Windows.Forms.ComboBox();
            this.cmbFolderPath = new System.Windows.Forms.ComboBox();
            this.btnSearch = new System.Windows.Forms.Button();
            this.chkUseRegex = new System.Windows.Forms.CheckBox();
            this.dataGridViewResults = new System.Windows.Forms.DataGridView();
            this.columnFilePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chkIncludeFolderNames = new System.Windows.Forms.CheckBox();
            this.labelResult = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResults)).BeginInit();
            this.SuspendLayout();
            // 
            // btnBrowse
            // 
            this.btnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowse.Location = new System.Drawing.Point(686, 28);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(45, 23);
            this.btnBrowse.TabIndex = 0;
            this.btnBrowse.Text = "ref";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(46, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "Search dir path";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(46, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(91, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "Search file name";
            // 
            // cmbKeyword
            // 
            this.cmbKeyword.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbKeyword.FormattingEnabled = true;
            this.cmbKeyword.Location = new System.Drawing.Point(48, 101);
            this.cmbKeyword.Name = "cmbKeyword";
            this.cmbKeyword.Size = new System.Drawing.Size(632, 20);
            this.cmbKeyword.TabIndex = 0;
            // 
            // cmbFolderPath
            // 
            this.cmbFolderPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFolderPath.FormattingEnabled = true;
            this.cmbFolderPath.Location = new System.Drawing.Point(48, 28);
            this.cmbFolderPath.Name = "cmbFolderPath";
            this.cmbFolderPath.Size = new System.Drawing.Size(632, 20);
            this.cmbFolderPath.TabIndex = 1;
            // 
            // btnSearch
            // 
            this.btnSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSearch.Location = new System.Drawing.Point(307, 132);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(75, 33);
            this.btnSearch.TabIndex = 5;
            this.btnSearch.Text = "SEARCH";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // chkUseRegex
            // 
            this.chkUseRegex.AutoSize = true;
            this.chkUseRegex.Location = new System.Drawing.Point(169, 81);
            this.chkUseRegex.Name = "chkUseRegex";
            this.chkUseRegex.Size = new System.Drawing.Size(88, 16);
            this.chkUseRegex.TabIndex = 6;
            this.chkUseRegex.Text = "RegEx Mode";
            this.chkUseRegex.UseVisualStyleBackColor = true;
            // 
            // dataGridViewResults
            // 
            this.dataGridViewResults.AllowUserToAddRows = false;
            this.dataGridViewResults.AllowUserToDeleteRows = false;
            this.dataGridViewResults.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewResults.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.columnFilePath});
            this.dataGridViewResults.Location = new System.Drawing.Point(48, 174);
            this.dataGridViewResults.Name = "dataGridViewResults";
            this.dataGridViewResults.ReadOnly = true;
            this.dataGridViewResults.Size = new System.Drawing.Size(660, 324);
            this.dataGridViewResults.TabIndex = 6;
            this.dataGridViewResults.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewResults_CellDoubleClick);
            // 
            // columnFilePath
            // 
            this.columnFilePath.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.columnFilePath.HeaderText = "ファイルパス";
            this.columnFilePath.Name = "columnFilePath";
            this.columnFilePath.ReadOnly = true;
            // 
            // chkIncludeFolderNames
            // 
            this.chkIncludeFolderNames.AutoSize = true;
            this.chkIncludeFolderNames.Location = new System.Drawing.Point(294, 81);
            this.chkIncludeFolderNames.Name = "chkIncludeFolderNames";
            this.chkIncludeFolderNames.Size = new System.Drawing.Size(104, 16);
            this.chkIncludeFolderNames.TabIndex = 7;
            this.chkIncludeFolderNames.Text = "With Dir Search";
            this.chkIncludeFolderNames.UseVisualStyleBackColor = true;
            // 
            // labelResult
            // 
            this.labelResult.AutoSize = true;
            this.labelResult.Location = new System.Drawing.Point(398, 142);
            this.labelResult.Name = "labelResult";
            this.labelResult.Size = new System.Drawing.Size(38, 12);
            this.labelResult.TabIndex = 8;
            this.labelResult.Text = "Result";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 510);
            this.Controls.Add(this.labelResult);
            this.Controls.Add(this.chkIncludeFolderNames);
            this.Controls.Add(this.dataGridViewResults);
            this.Controls.Add(this.chkUseRegex);
            this.Controls.Add(this.btnSearch);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbFolderPath);
            this.Controls.Add(this.cmbKeyword);
            this.Controls.Add(this.btnBrowse);
            this.Name = "MainForm";
            this.Text = "Simple File Search";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResults)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbKeyword;
        private System.Windows.Forms.ComboBox cmbFolderPath;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.CheckBox chkUseRegex;
        private System.Windows.Forms.DataGridView dataGridViewResults;
        private System.Windows.Forms.DataGridViewTextBoxColumn columnFilePath;
        private System.Windows.Forms.CheckBox chkIncludeFolderNames;
        private System.Windows.Forms.Label labelResult;
    }
}

