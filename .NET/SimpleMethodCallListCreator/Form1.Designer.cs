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
            this.clmFilePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmFileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmClassName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmCallerMethod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmCalleeClass = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmCalleeMethod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmCalleeMethodParams = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmRowNumCalleeMethod = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btnBrowse = new System.Windows.Forms.Button();
            this.cmbFilePath = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnRun = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.cmbCallerMethod = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtCallerMethodNameFilter = new System.Windows.Forms.TextBox();
            this.txtCalleeClassNameFilter = new System.Windows.Forms.TextBox();
            this.txtCalleeMethodNameFitter = new System.Windows.Forms.TextBox();
            this.txtCalleeMethodParamFilter = new System.Windows.Forms.TextBox();
            this.txtRowNumFilter = new System.Windows.Forms.TextBox();
            this.txtClassNameFilter = new System.Windows.Forms.TextBox();
            this.txtFileNameFilter = new System.Windows.Forms.TextBox();
            this.txtFilePathFilter = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
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
            this.dataGridViewResults.Location = new System.Drawing.Point(21, 206);
            this.dataGridViewResults.Name = "dataGridViewResults";
            this.dataGridViewResults.ReadOnly = true;
            this.dataGridViewResults.Size = new System.Drawing.Size(771, 241);
            this.dataGridViewResults.TabIndex = 16;
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
            // btnBrowse
            // 
            this.btnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowse.Location = new System.Drawing.Point(723, 28);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(45, 23);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "参照";
            this.btnBrowse.UseVisualStyleBackColor = true;
            // 
            // cmbFilePath
            // 
            this.cmbFilePath.AllowDrop = true;
            this.cmbFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFilePath.FormattingEnabled = true;
            this.cmbFilePath.Location = new System.Drawing.Point(48, 28);
            this.cmbFilePath.Name = "cmbFilePath";
            this.cmbFilePath.Size = new System.Drawing.Size(669, 20);
            this.cmbFilePath.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(46, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "対象ファイルパス";
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(307, 117);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 33);
            this.btnRun.TabIndex = 5;
            this.btnRun.Text = "実行";
            this.btnRun.UseVisualStyleBackColor = true;
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(584, 123);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 21);
            this.btnExport.TabIndex = 6;
            this.btnExport.Text = "結果出力";
            this.btnExport.UseVisualStyleBackColor = true;
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(665, 123);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 21);
            this.btnImport.TabIndex = 7;
            this.btnImport.Text = "結果入力";
            this.btnImport.UseVisualStyleBackColor = true;
            // 
            // cmbCallerMethod
            // 
            this.cmbCallerMethod.AllowDrop = true;
            this.cmbCallerMethod.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbCallerMethod.FormattingEnabled = true;
            this.cmbCallerMethod.Location = new System.Drawing.Point(48, 76);
            this.cmbCallerMethod.Name = "cmbCallerMethod";
            this.cmbCallerMethod.Size = new System.Drawing.Size(669, 20);
            this.cmbCallerMethod.TabIndex = 3;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(46, 61);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(74, 12);
            this.label4.TabIndex = 0;
            this.label4.Text = "呼出元メソッド";
            // 
            // txtCallerMethodNameFilter
            // 
            this.txtCallerMethodNameFilter.Location = new System.Drawing.Point(285, 181);
            this.txtCallerMethodNameFilter.Name = "txtCallerMethodNameFilter";
            this.txtCallerMethodNameFilter.Size = new System.Drawing.Size(131, 19);
            this.txtCallerMethodNameFilter.TabIndex = 11;
            // 
            // txtCalleeClassNameFilter
            // 
            this.txtCalleeClassNameFilter.Location = new System.Drawing.Point(422, 181);
            this.txtCalleeClassNameFilter.Name = "txtCalleeClassNameFilter";
            this.txtCalleeClassNameFilter.Size = new System.Drawing.Size(84, 19);
            this.txtCalleeClassNameFilter.TabIndex = 12;
            // 
            // txtCalleeMethodNameFitter
            // 
            this.txtCalleeMethodNameFitter.Location = new System.Drawing.Point(512, 181);
            this.txtCalleeMethodNameFitter.Name = "txtCalleeMethodNameFitter";
            this.txtCalleeMethodNameFitter.Size = new System.Drawing.Size(110, 19);
            this.txtCalleeMethodNameFitter.TabIndex = 13;
            // 
            // txtCalleeMethodParamFilter
            // 
            this.txtCalleeMethodParamFilter.Location = new System.Drawing.Point(628, 181);
            this.txtCalleeMethodParamFilter.Name = "txtCalleeMethodParamFilter";
            this.txtCalleeMethodParamFilter.Size = new System.Drawing.Size(76, 19);
            this.txtCalleeMethodParamFilter.TabIndex = 14;
            // 
            // txtRowNumFilter
            // 
            this.txtRowNumFilter.Location = new System.Drawing.Point(710, 181);
            this.txtRowNumFilter.Name = "txtRowNumFilter";
            this.txtRowNumFilter.Size = new System.Drawing.Size(49, 19);
            this.txtRowNumFilter.TabIndex = 15;
            // 
            // txtClassNameFilter
            // 
            this.txtClassNameFilter.Location = new System.Drawing.Point(195, 181);
            this.txtClassNameFilter.Name = "txtClassNameFilter";
            this.txtClassNameFilter.Size = new System.Drawing.Size(84, 19);
            this.txtClassNameFilter.TabIndex = 10;
            // 
            // txtFileNameFilter
            // 
            this.txtFileNameFilter.Location = new System.Drawing.Point(108, 181);
            this.txtFileNameFilter.Name = "txtFileNameFilter";
            this.txtFileNameFilter.Size = new System.Drawing.Size(84, 19);
            this.txtFileNameFilter.TabIndex = 9;
            // 
            // txtFilePathFilter
            // 
            this.txtFilePathFilter.Location = new System.Drawing.Point(18, 181);
            this.txtFilePathFilter.Name = "txtFilePathFilter";
            this.txtFilePathFilter.Size = new System.Drawing.Size(84, 19);
            this.txtFilePathFilter.TabIndex = 8;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(16, 166);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(58, 12);
            this.label5.TabIndex = 0;
            this.label5.Text = "ファイルパス";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(106, 166);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(51, 12);
            this.label6.TabIndex = 0;
            this.label6.Text = "ファイル名";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(196, 166);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(30, 12);
            this.label7.TabIndex = 0;
            this.label7.Text = "クラス";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(283, 166);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(74, 12);
            this.label8.TabIndex = 0;
            this.label8.Text = "呼出元メソッド";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(420, 166);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(66, 12);
            this.label9.TabIndex = 0;
            this.label9.Text = "呼出先クラス";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(510, 166);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(74, 12);
            this.label10.TabIndex = 0;
            this.label10.Text = "呼出先メソッド";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(626, 166);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(29, 12);
            this.label11.TabIndex = 0;
            this.label11.Text = "引数";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(708, 166);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(41, 12);
            this.label12.TabIndex = 0;
            this.label12.Text = "行番号";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(48, 117);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(80, 24);
            this.button1.TabIndex = 4;
            this.button1.Text = "除外指定";
            this.toolTip1.SetToolTip(this.button1, "呼出先メソッド名の除外指定");
            this.button1.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(804, 459);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtFilePathFilter);
            this.Controls.Add(this.txtFileNameFilter);
            this.Controls.Add(this.txtClassNameFilter);
            this.Controls.Add(this.txtRowNumFilter);
            this.Controls.Add(this.txtCalleeMethodParamFilter);
            this.Controls.Add(this.txtCalleeMethodNameFitter);
            this.Controls.Add(this.txtCalleeClassNameFilter);
            this.Controls.Add(this.txtCallerMethodNameFilter);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cmbCallerMethod);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.dataGridViewResults);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbFilePath);
            this.Controls.Add(this.btnBrowse);
            this.Name = "MainForm";
            this.Text = "Simple Method CallList Creator";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResults)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.DataGridView dataGridViewResults;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.ComboBox cmbFilePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmFilePath;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmFileName;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmClassName;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmCallerMethod;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmCalleeClass;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmCalleeMethod;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmCalleeMethodParams;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmRowNumCalleeMethod;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.ComboBox cmbCallerMethod;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtCallerMethodNameFilter;
        private System.Windows.Forms.TextBox txtCalleeClassNameFilter;
        private System.Windows.Forms.TextBox txtCalleeMethodNameFitter;
        private System.Windows.Forms.TextBox txtCalleeMethodParamFilter;
        private System.Windows.Forms.TextBox txtRowNumFilter;
        private System.Windows.Forms.TextBox txtClassNameFilter;
        private System.Windows.Forms.TextBox txtFileNameFilter;
        private System.Windows.Forms.TextBox txtFilePathFilter;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Button button1;
    }
}

