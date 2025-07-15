namespace SimpleGrep
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
            this.btnCancel = new System.Windows.Forms.Button();
            this.chkUseRegex = new System.Windows.Forms.CheckBox();
            this.dataGridViewResults = new System.Windows.Forms.DataGridView();
            this.clmFilePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmLine = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmGrepResult = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.labelResult = new System.Windows.Forms.Label();
            this.chkSearchSubDir = new System.Windows.Forms.CheckBox();
            this.chkCase = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.btnExportSakura = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.chkTagJump = new System.Windows.Forms.CheckBox();
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
            this.btnBrowse.Text = "参照";
            this.btnBrowse.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(46, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "検索フォルダパス";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(46, 141);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "検索キーワード";
            // 
            // cmbKeyword
            // 
            this.cmbKeyword.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbKeyword.FormattingEnabled = true;
            this.cmbKeyword.Location = new System.Drawing.Point(48, 156);
            this.cmbKeyword.Name = "cmbKeyword";
            this.cmbKeyword.Size = new System.Drawing.Size(632, 20);
            this.cmbKeyword.TabIndex = 0;
            // 
            // cmbFolderPath
            // 
            this.cmbFolderPath.AllowDrop = true;
            this.cmbFolderPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFolderPath.FormattingEnabled = true;
            this.cmbFolderPath.Location = new System.Drawing.Point(48, 28);
            this.cmbFolderPath.Name = "cmbFolderPath";
            this.cmbFolderPath.Size = new System.Drawing.Size(632, 20);
            this.cmbFolderPath.TabIndex = 1;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(352, 212);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 33);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "中止";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // chkUseRegex
            // 
            this.chkUseRegex.AutoSize = true;
            this.chkUseRegex.Location = new System.Drawing.Point(51, 182);
            this.chkUseRegex.Name = "chkUseRegex";
            this.chkUseRegex.Size = new System.Drawing.Size(72, 16);
            this.chkUseRegex.TabIndex = 6;
            this.chkUseRegex.Text = "正規表現";
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
            this.clmFilePath,
            this.clmLine,
            this.clmGrepResult});
            this.dataGridViewResults.Location = new System.Drawing.Point(48, 284);
            this.dataGridViewResults.Name = "dataGridViewResults";
            this.dataGridViewResults.ReadOnly = true;
            this.dataGridViewResults.Size = new System.Drawing.Size(660, 311);
            this.dataGridViewResults.TabIndex = 6;
            // 
            // clmFilePath
            // 
            this.clmFilePath.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmFilePath.HeaderText = "ファイルパス";
            this.clmFilePath.Name = "clmFilePath";
            this.clmFilePath.ReadOnly = true;
            // 
            // clmLine
            // 
            this.clmLine.HeaderText = "行";
            this.clmLine.Name = "clmLine";
            this.clmLine.ReadOnly = true;
            // 
            // clmGrepResult
            // 
            this.clmGrepResult.HeaderText = "Grep結果";
            this.clmGrepResult.Name = "clmGrepResult";
            this.clmGrepResult.ReadOnly = true;
            // 
            // labelResult
            // 
            this.labelResult.AutoSize = true;
            this.labelResult.Location = new System.Drawing.Point(443, 222);
            this.labelResult.Name = "labelResult";
            this.labelResult.Size = new System.Drawing.Size(38, 12);
            this.labelResult.TabIndex = 8;
            this.labelResult.Text = "Result";
            // 
            // chkSearchSubDir
            // 
            this.chkSearchSubDir.AutoSize = true;
            this.chkSearchSubDir.Location = new System.Drawing.Point(48, 54);
            this.chkSearchSubDir.Name = "chkSearchSubDir";
            this.chkSearchSubDir.Size = new System.Drawing.Size(111, 16);
            this.chkSearchSubDir.TabIndex = 68;
            this.chkSearchSubDir.Text = "サブフォルダも対象";
            this.chkSearchSubDir.UseVisualStyleBackColor = true;
            // 
            // chkCase
            // 
            this.chkCase.AutoSize = true;
            this.chkCase.Location = new System.Drawing.Point(141, 182);
            this.chkCase.Name = "chkCase";
            this.chkCase.Size = new System.Drawing.Size(72, 16);
            this.chkCase.TabIndex = 69;
            this.chkCase.Text = "大小区別";
            this.chkCase.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(271, 212);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 33);
            this.button1.TabIndex = 70;
            this.button1.Text = "検索";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // btnExportSakura
            // 
            this.btnExportSakura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportSakura.Location = new System.Drawing.Point(48, 222);
            this.btnExportSakura.Name = "btnExportSakura";
            this.btnExportSakura.Size = new System.Drawing.Size(96, 23);
            this.btnExportSakura.TabIndex = 71;
            this.btnExportSakura.Text = "Excport sakura";
            this.btnExportSakura.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(49, 83);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(113, 12);
            this.label3.TabIndex = 72;
            this.label3.Text = "対象ファイル（例：*.cs）";
            // 
            // comboBox1
            // 
            this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(48, 98);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(632, 20);
            this.comboBox1.TabIndex = 73;
            // 
            // chkTagJump
            // 
            this.chkTagJump.AutoSize = true;
            this.chkTagJump.Location = new System.Drawing.Point(233, 182);
            this.chkTagJump.Name = "chkTagJump";
            this.chkTagJump.Size = new System.Drawing.Size(77, 16);
            this.chkTagJump.TabIndex = 74;
            this.chkTagJump.Text = "タグジャンプ";
            this.chkTagJump.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 607);
            this.Controls.Add(this.chkTagJump);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnExportSakura);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.chkCase);
            this.Controls.Add(this.chkSearchSubDir);
            this.Controls.Add(this.labelResult);
            this.Controls.Add(this.dataGridViewResults);
            this.Controls.Add(this.chkUseRegex);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbFolderPath);
            this.Controls.Add(this.cmbKeyword);
            this.Controls.Add(this.btnBrowse);
            this.Name = "MainForm";
            this.Text = "Simple Grep";
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
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.CheckBox chkUseRegex;
        private System.Windows.Forms.DataGridView dataGridViewResults;
        private System.Windows.Forms.Label labelResult;
        private System.Windows.Forms.CheckBox chkSearchSubDir;
        private System.Windows.Forms.CheckBox chkCase;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmFilePath;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmLine;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmGrepResult;
        private System.Windows.Forms.Button btnExportSakura;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.CheckBox chkTagJump;
    }
}

