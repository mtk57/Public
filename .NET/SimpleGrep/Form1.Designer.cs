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
            this.components = new System.ComponentModel.Container();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbKeyword = new System.Windows.Forms.ComboBox();
            this.cmbFolderPath = new System.Windows.Forms.ComboBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.chkUseRegex = new System.Windows.Forms.CheckBox();
            this.dataGridViewResults = new System.Windows.Forms.DataGridView();
            this.clmFilePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmFileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmExtension = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmLine = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmGrepResult = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmMethodSignature = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.labelTime = new System.Windows.Forms.Label();
            this.chkSearchSubDir = new System.Windows.Forms.CheckBox();
            this.chkCase = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.btnExportSakura = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.chkTagJump = new System.Windows.Forms.CheckBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblPer = new System.Windows.Forms.Label();
            this.chkMethod = new System.Windows.Forms.CheckBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.chkIgnoreComment = new System.Windows.Forms.CheckBox();
            this.btnFileCopy = new System.Windows.Forms.Button();
            this.txtFilePathFilter = new System.Windows.Forms.TextBox();
            this.txtFileNameFilter = new System.Windows.Forms.TextBox();
            this.txtExtensionFilter = new System.Windows.Forms.TextBox();
            this.txtRowNumFilter = new System.Windows.Forms.TextBox();
            this.txtGrepResultFilter = new System.Windows.Forms.TextBox();
            this.txtMethodFilter = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.btnMultiKeywords = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.chkIgnoreBinaryFile = new System.Windows.Forms.CheckBox();
            this.cmbExcludeFolder = new System.Windows.Forms.ComboBox();
            this.cmbExcludeExtension = new System.Windows.Forms.ComboBox();
            this.lblResultCount = new System.Windows.Forms.Label();
            this.btnCollectExtensions = new System.Windows.Forms.Button();
            this.btnClearFilters = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResults)).BeginInit();
            this.SuspendLayout();
            // 
            // btnBrowse
            // 
            this.btnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowse.Location = new System.Drawing.Point(696, 28);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(45, 23);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "参照";
            this.btnBrowse.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(46, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "検索フォルダパス";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(46, 131);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 0;
            this.label2.Text = "検索キーワード";
            // 
            // cmbKeyword
            // 
            this.cmbKeyword.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbKeyword.FormattingEnabled = true;
            this.cmbKeyword.Location = new System.Drawing.Point(48, 146);
            this.cmbKeyword.Name = "cmbKeyword";
            this.cmbKeyword.Size = new System.Drawing.Size(642, 20);
            this.cmbKeyword.TabIndex = 8;
            // 
            // cmbFolderPath
            // 
            this.cmbFolderPath.AllowDrop = true;
            this.cmbFolderPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFolderPath.FormattingEnabled = true;
            this.cmbFolderPath.Location = new System.Drawing.Point(48, 28);
            this.cmbFolderPath.Name = "cmbFolderPath";
            this.cmbFolderPath.Size = new System.Drawing.Size(642, 20);
            this.cmbFolderPath.TabIndex = 1;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(352, 202);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 33);
            this.btnCancel.TabIndex = 17;
            this.btnCancel.Text = "中止";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // chkUseRegex
            // 
            this.chkUseRegex.AutoSize = true;
            this.chkUseRegex.Location = new System.Drawing.Point(51, 172);
            this.chkUseRegex.Name = "chkUseRegex";
            this.chkUseRegex.Size = new System.Drawing.Size(72, 16);
            this.chkUseRegex.TabIndex = 9;
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
            this.clmFileName,
            this.clmExtension,
            this.clmLine,
            this.clmGrepResult,
            this.clmMethodSignature});
            this.dataGridViewResults.Location = new System.Drawing.Point(12, 305);
            this.dataGridViewResults.Name = "dataGridViewResults";
            this.dataGridViewResults.ReadOnly = true;
            this.dataGridViewResults.Size = new System.Drawing.Size(753, 138);
            this.dataGridViewResults.TabIndex = 25;
            // 
            // clmFilePath
            // 
            this.clmFilePath.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmFilePath.HeaderText = "ファイルパス";
            this.clmFilePath.Name = "clmFilePath";
            this.clmFilePath.ReadOnly = true;
            // 
            // clmFileName
            // 
            this.clmFileName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmFileName.HeaderText = "ファイル名";
            this.clmFileName.Name = "clmFileName";
            this.clmFileName.ReadOnly = true;
            // 
            // clmExtension
            // 
            this.clmExtension.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmExtension.HeaderText = "拡張子";
            this.clmExtension.Name = "clmExtension";
            this.clmExtension.ReadOnly = true;
            // 
            // clmLine
            // 
            this.clmLine.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmLine.HeaderText = "行";
            this.clmLine.Name = "clmLine";
            this.clmLine.ReadOnly = true;
            // 
            // clmGrepResult
            // 
            this.clmGrepResult.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmGrepResult.HeaderText = "Grep結果";
            this.clmGrepResult.Name = "clmGrepResult";
            this.clmGrepResult.ReadOnly = true;
            // 
            // clmMethodSignature
            // 
            this.clmMethodSignature.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmMethodSignature.HeaderText = "メソッド";
            this.clmMethodSignature.Name = "clmMethodSignature";
            this.clmMethodSignature.ReadOnly = true;
            // 
            // labelTime
            // 
            this.labelTime.AutoSize = true;
            this.labelTime.Location = new System.Drawing.Point(433, 212);
            this.labelTime.Name = "labelTime";
            this.labelTime.Size = new System.Drawing.Size(30, 12);
            this.labelTime.TabIndex = 0;
            this.labelTime.Text = "Time";
            // 
            // chkSearchSubDir
            // 
            this.chkSearchSubDir.AutoSize = true;
            this.chkSearchSubDir.Location = new System.Drawing.Point(48, 54);
            this.chkSearchSubDir.Name = "chkSearchSubDir";
            this.chkSearchSubDir.Size = new System.Drawing.Size(111, 16);
            this.chkSearchSubDir.TabIndex = 3;
            this.chkSearchSubDir.Text = "サブフォルダも対象";
            this.chkSearchSubDir.UseVisualStyleBackColor = true;
            // 
            // chkCase
            // 
            this.chkCase.AutoSize = true;
            this.chkCase.Location = new System.Drawing.Point(141, 172);
            this.chkCase.Name = "chkCase";
            this.chkCase.Size = new System.Drawing.Size(72, 16);
            this.chkCase.TabIndex = 10;
            this.chkCase.Text = "大小区別";
            this.chkCase.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(271, 202);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 33);
            this.button1.TabIndex = 16;
            this.button1.Text = "検索";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // btnExportSakura
            // 
            this.btnExportSakura.Location = new System.Drawing.Point(51, 202);
            this.btnExportSakura.Name = "btnExportSakura";
            this.btnExportSakura.Size = new System.Drawing.Size(96, 23);
            this.btnExportSakura.TabIndex = 15;
            this.btnExportSakura.Text = "Export sakura";
            this.btnExportSakura.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(49, 83);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(113, 12);
            this.label3.TabIndex = 0;
            this.label3.Text = "対象ファイル（例：*.cs）";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(48, 98);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(167, 20);
            this.comboBox1.TabIndex = 4;
            // 
            // chkTagJump
            // 
            this.chkTagJump.AutoSize = true;
            this.chkTagJump.Location = new System.Drawing.Point(53, 231);
            this.chkTagJump.Name = "chkTagJump";
            this.chkTagJump.Size = new System.Drawing.Size(77, 16);
            this.chkTagJump.TabIndex = 11;
            this.chkTagJump.Text = "タグジャンプ";
            this.chkTagJump.UseVisualStyleBackColor = true;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(271, 241);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(156, 10);
            this.progressBar.TabIndex = 15;
            // 
            // lblPer
            // 
            this.lblPer.AutoSize = true;
            this.lblPer.Location = new System.Drawing.Point(433, 239);
            this.lblPer.Name = "lblPer";
            this.lblPer.Size = new System.Drawing.Size(21, 12);
            this.lblPer.TabIndex = 0;
            this.lblPer.Text = "0 %";
            // 
            // chkMethod
            // 
            this.chkMethod.AutoSize = true;
            this.chkMethod.Location = new System.Drawing.Point(231, 172);
            this.chkMethod.Name = "chkMethod";
            this.chkMethod.Size = new System.Drawing.Size(102, 16);
            this.chkMethod.TabIndex = 12;
            this.chkMethod.Text = "メソッド名を導出";
            this.toolTip1.SetToolTip(this.chkMethod, "現バージョンではJavaのみをサポート");
            this.chkMethod.UseVisualStyleBackColor = true;
            // 
            // chkIgnoreComment
            // 
            this.chkIgnoreComment.AutoSize = true;
            this.chkIgnoreComment.Location = new System.Drawing.Point(351, 172);
            this.chkIgnoreComment.Name = "chkIgnoreComment";
            this.chkIgnoreComment.Size = new System.Drawing.Size(81, 16);
            this.chkIgnoreComment.TabIndex = 13;
            this.chkIgnoreComment.Text = "コメント無視";
            this.toolTip1.SetToolTip(this.chkIgnoreComment, "現バージョンではJavaのみをサポート");
            this.chkIgnoreComment.UseVisualStyleBackColor = true;
            // 
            // btnFileCopy
            // 
            this.btnFileCopy.Location = new System.Drawing.Point(605, 202);
            this.btnFileCopy.Name = "btnFileCopy";
            this.btnFileCopy.Size = new System.Drawing.Size(75, 23);
            this.btnFileCopy.TabIndex = 18;
            this.btnFileCopy.Text = "ファイルコピー";
            this.btnFileCopy.UseVisualStyleBackColor = true;
            // 
            // txtFilePathFilter
            // 
            this.txtFilePathFilter.Location = new System.Drawing.Point(11, 280);
            this.txtFilePathFilter.Name = "txtFilePathFilter";
            this.txtFilePathFilter.Size = new System.Drawing.Size(81, 19);
            this.txtFilePathFilter.TabIndex = 20;
            // 
            // txtFileNameFilter
            // 
            this.txtFileNameFilter.Location = new System.Drawing.Point(98, 280);
            this.txtFileNameFilter.Name = "txtFileNameFilter";
            this.txtFileNameFilter.Size = new System.Drawing.Size(81, 19);
            this.txtFileNameFilter.TabIndex = 21;
            // 
            // txtExtensionFilter
            // 
            this.txtExtensionFilter.Location = new System.Drawing.Point(185, 280);
            this.txtExtensionFilter.Name = "txtExtensionFilter";
            this.txtExtensionFilter.Size = new System.Drawing.Size(81, 19);
            this.txtExtensionFilter.TabIndex = 22;
            // 
            // txtRowNumFilter
            // 
            this.txtRowNumFilter.Location = new System.Drawing.Point(272, 280);
            this.txtRowNumFilter.Name = "txtRowNumFilter";
            this.txtRowNumFilter.Size = new System.Drawing.Size(81, 19);
            this.txtRowNumFilter.TabIndex = 23;
            // 
            // txtGrepResultFilter
            // 
            this.txtGrepResultFilter.Location = new System.Drawing.Point(359, 280);
            this.txtGrepResultFilter.Name = "txtGrepResultFilter";
            this.txtGrepResultFilter.Size = new System.Drawing.Size(81, 19);
            this.txtGrepResultFilter.TabIndex = 24;
            // 
            // txtMethodFilter
            // 
            this.txtMethodFilter.Location = new System.Drawing.Point(446, 280);
            this.txtMethodFilter.Name = "txtMethodFilter";
            this.txtMethodFilter.Size = new System.Drawing.Size(81, 19);
            this.txtMethodFilter.TabIndex = 25;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 265);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(58, 12);
            this.label4.TabIndex = 0;
            this.label4.Text = "ファイルパス";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(102, 265);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(51, 12);
            this.label5.TabIndex = 0;
            this.label5.Text = "ファイル名";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(281, 265);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(17, 12);
            this.label6.TabIndex = 0;
            this.label6.Text = "行";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(357, 265);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 12);
            this.label7.TabIndex = 0;
            this.label7.Text = "Grep結果";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(444, 265);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(38, 12);
            this.label8.TabIndex = 0;
            this.label8.Text = "メソッド";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(184, 265);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(41, 12);
            this.label11.TabIndex = 0;
            this.label11.Text = "拡張子";
            // 
            // btnMultiKeywords
            // 
            this.btnMultiKeywords.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMultiKeywords.Location = new System.Drawing.Point(699, 146);
            this.btnMultiKeywords.Name = "btnMultiKeywords";
            this.btnMultiKeywords.Size = new System.Drawing.Size(45, 20);
            this.btnMultiKeywords.TabIndex = 26;
            this.btnMultiKeywords.Text = "複数";
            this.btnMultiKeywords.UseVisualStyleBackColor = true;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(252, 83);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(176, 12);
            this.label9.TabIndex = 0;
            this.label9.Text = "除外フォルダ（/区切り、例: bin/obj）";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(459, 83);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(177, 12);
            this.label10.TabIndex = 0;
            this.label10.Text = "除外拡張子（/区切り、例: exe/dll）";
            // 
            // chkIgnoreBinaryFile
            // 
            this.chkIgnoreBinaryFile.AutoSize = true;
            this.chkIgnoreBinaryFile.Location = new System.Drawing.Point(461, 172);
            this.chkIgnoreBinaryFile.Name = "chkIgnoreBinaryFile";
            this.chkIgnoreBinaryFile.Size = new System.Drawing.Size(127, 16);
            this.chkIgnoreBinaryFile.TabIndex = 14;
            this.chkIgnoreBinaryFile.Text = "バイナリファイルを無視";
            this.chkIgnoreBinaryFile.UseVisualStyleBackColor = true;
            // 
            // cmbExcludeFolder
            // 
            this.cmbExcludeFolder.FormattingEnabled = true;
            this.cmbExcludeFolder.Location = new System.Drawing.Point(254, 98);
            this.cmbExcludeFolder.Name = "cmbExcludeFolder";
            this.cmbExcludeFolder.Size = new System.Drawing.Size(177, 20);
            this.cmbExcludeFolder.TabIndex = 5;
            // 
            // cmbExcludeExtension
            // 
            this.cmbExcludeExtension.FormattingEnabled = true;
            this.cmbExcludeExtension.Location = new System.Drawing.Point(461, 98);
            this.cmbExcludeExtension.Name = "cmbExcludeExtension";
            this.cmbExcludeExtension.Size = new System.Drawing.Size(195, 20);
            this.cmbExcludeExtension.TabIndex = 6;
            // 
            // lblResultCount
            // 
            this.lblResultCount.AutoSize = true;
            this.lblResultCount.Location = new System.Drawing.Point(201, 212);
            this.lblResultCount.Name = "lblResultCount";
            this.lblResultCount.Size = new System.Drawing.Size(0, 12);
            this.lblResultCount.TabIndex = 0;
            // 
            // btnCollectExtensions
            // 
            this.btnCollectExtensions.Location = new System.Drawing.Point(662, 96);
            this.btnCollectExtensions.Name = "btnCollectExtensions";
            this.btnCollectExtensions.Size = new System.Drawing.Size(102, 23);
            this.btnCollectExtensions.TabIndex = 27;
            this.btnCollectExtensions.Text = "全拡張子を収集";
            this.btnCollectExtensions.UseVisualStyleBackColor = true;
            // 
            // btnClearFilters
            // 
            this.btnClearFilters.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClearFilters.Location = new System.Drawing.Point(716, 281);
            this.btnClearFilters.Name = "btnClearFilters";
            this.btnClearFilters.Size = new System.Drawing.Size(45, 21);
            this.btnClearFilters.TabIndex = 28;
            this.btnClearFilters.Text = "クリア";
            this.btnClearFilters.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(777, 455);
            this.Controls.Add(this.btnCollectExtensions);
            this.Controls.Add(this.btnClearFilters);
            this.Controls.Add(this.lblResultCount);
            this.Controls.Add(this.chkIgnoreBinaryFile);
            this.Controls.Add(this.cmbExcludeExtension);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.cmbExcludeFolder);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.btnMultiKeywords);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtMethodFilter);
            this.Controls.Add(this.txtGrepResultFilter);
            this.Controls.Add(this.txtRowNumFilter);
            this.Controls.Add(this.txtExtensionFilter);
            this.Controls.Add(this.txtFileNameFilter);
            this.Controls.Add(this.txtFilePathFilter);
            this.Controls.Add(this.chkIgnoreComment);
            this.Controls.Add(this.btnFileCopy);
            this.Controls.Add(this.chkMethod);
            this.Controls.Add(this.lblPer);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.chkTagJump);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnExportSakura);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.chkCase);
            this.Controls.Add(this.chkSearchSubDir);
            this.Controls.Add(this.labelTime);
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
        private System.Windows.Forms.Label labelTime;
        private System.Windows.Forms.CheckBox chkSearchSubDir;
        private System.Windows.Forms.CheckBox chkCase;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnExportSakura;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.CheckBox chkTagJump;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblPer;
        private System.Windows.Forms.CheckBox chkMethod;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button btnFileCopy;
        private System.Windows.Forms.CheckBox chkIgnoreComment;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmFilePath;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmFileName;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmExtension;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmLine;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmGrepResult;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmMethodSignature;
        private System.Windows.Forms.TextBox txtFilePathFilter;
        private System.Windows.Forms.TextBox txtFileNameFilter;
        private System.Windows.Forms.TextBox txtExtensionFilter;
        private System.Windows.Forms.TextBox txtRowNumFilter;
        private System.Windows.Forms.TextBox txtGrepResultFilter;
        private System.Windows.Forms.TextBox txtMethodFilter;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button btnMultiKeywords;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.CheckBox chkIgnoreBinaryFile;
        private System.Windows.Forms.ComboBox cmbExcludeFolder;
        private System.Windows.Forms.ComboBox cmbExcludeExtension;
        private System.Windows.Forms.Label lblResultCount;
        private System.Windows.Forms.Button btnCollectExtensions;
        private System.Windows.Forms.Button btnClearFilters;
    }
}
