namespace SimpleMethodCallListCreator
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
            this.dataGridViewResults = new System.Windows.Forms.DataGridView();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btnBrowse = new System.Windows.Forms.Button();
            this.cmbIgnoreKeyword = new System.Windows.Forms.ComboBox();
            this.cmbFilePath = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.chkUseRegex = new System.Windows.Forms.CheckBox();
            this.chkCase = new System.Windows.Forms.CheckBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbIgnoreRules = new System.Windows.Forms.ComboBox();
            this.clmFilePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmFileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmClassName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmCallerMethod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmCalleeClass = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmCalleeMethod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmCalleeMethodParams = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmRowNumCalleeMethod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResults)).BeginInit();
            this.SuspendLayout();
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
            this.clmClassName,
            this.clmCallerMethod,
            this.clmCalleeClass,
            this.clmCalleeMethod,
            this.clmCalleeMethodParams,
            this.clmRowNumCalleeMethod});
            this.dataGridViewResults.Location = new System.Drawing.Point(48, 197);
            this.dataGridViewResults.Name = "dataGridViewResults";
            this.dataGridViewResults.ReadOnly = true;
            this.dataGridViewResults.Size = new System.Drawing.Size(660, 175);
            this.dataGridViewResults.TabIndex = 6;
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
            // cmbIgnoreKeyword
            // 
            this.cmbIgnoreKeyword.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbIgnoreKeyword.FormattingEnabled = true;
            this.cmbIgnoreKeyword.Location = new System.Drawing.Point(48, 83);
            this.cmbIgnoreKeyword.Name = "cmbIgnoreKeyword";
            this.cmbIgnoreKeyword.Size = new System.Drawing.Size(632, 20);
            this.cmbIgnoreKeyword.TabIndex = 0;
            // 
            // cmbFilePath
            // 
            this.cmbFilePath.AllowDrop = true;
            this.cmbFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFilePath.FormattingEnabled = true;
            this.cmbFilePath.Location = new System.Drawing.Point(48, 28);
            this.cmbFilePath.Name = "cmbFilePath";
            this.cmbFilePath.Size = new System.Drawing.Size(632, 20);
            this.cmbFilePath.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(46, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "対象ファイルパス";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(46, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "除外キーワード";
            // 
            // chkUseRegex
            // 
            this.chkUseRegex.AutoSize = true;
            this.chkUseRegex.Location = new System.Drawing.Point(48, 109);
            this.chkUseRegex.Name = "chkUseRegex";
            this.chkUseRegex.Size = new System.Drawing.Size(72, 16);
            this.chkUseRegex.TabIndex = 6;
            this.chkUseRegex.Text = "正規表現";
            this.chkUseRegex.UseVisualStyleBackColor = true;
            // 
            // chkCase
            // 
            this.chkCase.AutoSize = true;
            this.chkCase.Location = new System.Drawing.Point(144, 109);
            this.chkCase.Name = "chkCase";
            this.chkCase.Size = new System.Drawing.Size(72, 16);
            this.chkCase.TabIndex = 69;
            this.chkCase.Text = "大小区別";
            this.chkCase.UseVisualStyleBackColor = true;
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(298, 148);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 33);
            this.btnRun.TabIndex = 70;
            this.btnRun.Text = "実行";
            this.btnRun.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(233, 110);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 12);
            this.label3.TabIndex = 79;
            this.label3.Text = "除外ルール";
            // 
            // cmbIgnoreRules
            // 
            this.cmbIgnoreRules.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbIgnoreRules.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbIgnoreRules.FormattingEnabled = true;
            this.cmbIgnoreRules.Items.AddRange(new object[] {
            "始まる",
            "終わる",
            "含む"});
            this.cmbIgnoreRules.Location = new System.Drawing.Point(298, 107);
            this.cmbIgnoreRules.Name = "cmbIgnoreRules";
            this.cmbIgnoreRules.Size = new System.Drawing.Size(84, 20);
            this.cmbIgnoreRules.TabIndex = 80;
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
            // clmClassName
            // 
            this.clmClassName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmClassName.HeaderText = "クラス";
            this.clmClassName.Name = "clmClassName";
            this.clmClassName.ReadOnly = true;
            // 
            // clmCallerMethod
            // 
            this.clmCallerMethod.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmCallerMethod.HeaderText = "呼出元メソッド";
            this.clmCallerMethod.Name = "clmCallerMethod";
            this.clmCallerMethod.ReadOnly = true;
            // 
            // clmCalleeClass
            // 
            this.clmCalleeClass.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmCalleeClass.HeaderText = "呼出先クラス";
            this.clmCalleeClass.Name = "clmCalleeClass";
            this.clmCalleeClass.ReadOnly = true;
            // 
            // clmCalleeMethod
            // 
            this.clmCalleeMethod.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmCalleeMethod.HeaderText = "呼出先メソッド";
            this.clmCalleeMethod.Name = "clmCalleeMethod";
            this.clmCalleeMethod.ReadOnly = true;
            // 
            // clmCalleeMethodParams
            // 
            this.clmCalleeMethodParams.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmCalleeMethodParams.HeaderText = "呼出先メソッド引数";
            this.clmCalleeMethodParams.Name = "clmCalleeMethodParams";
            this.clmCalleeMethodParams.ReadOnly = true;
            // 
            // clmRowNumCalleeMethod
            // 
            this.clmRowNumCalleeMethod.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmRowNumCalleeMethod.HeaderText = "行番号";
            this.clmRowNumCalleeMethod.Name = "clmRowNumCalleeMethod";
            this.clmRowNumCalleeMethod.ReadOnly = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 400);
            this.Controls.Add(this.cmbIgnoreRules);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.chkCase);
            this.Controls.Add(this.dataGridViewResults);
            this.Controls.Add(this.chkUseRegex);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbFilePath);
            this.Controls.Add(this.cmbIgnoreKeyword);
            this.Controls.Add(this.btnBrowse);
            this.Name = "MainForm";
            this.Text = "Simple Grep";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResults)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.DataGridView dataGridViewResults;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.ComboBox cmbIgnoreKeyword;
        private System.Windows.Forms.ComboBox cmbFilePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox chkUseRegex;
        private System.Windows.Forms.CheckBox chkCase;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbIgnoreRules;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmFilePath;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmFileName;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmClassName;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmCallerMethod;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmCalleeClass;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmCalleeMethod;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmCalleeMethodParams;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmRowNumCalleeMethod;
    }
}

