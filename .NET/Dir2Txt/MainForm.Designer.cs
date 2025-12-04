namespace Dir2Txt
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
            this.txtDirPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnRefDirPath = new System.Windows.Forms.Button();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.btnExtract = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtIgnoreDirs = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtIgnoreFiles = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // txtDirPath
            // 
            this.txtDirPath.Location = new System.Drawing.Point(22, 24);
            this.txtDirPath.Name = "txtDirPath";
            this.txtDirPath.Size = new System.Drawing.Size(400, 19);
            this.txtDirPath.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "フォルダパス";
            // 
            // btnRefDirPath
            // 
            this.btnRefDirPath.Location = new System.Drawing.Point(429, 19);
            this.btnRefDirPath.Name = "btnRefDirPath";
            this.btnRefDirPath.Size = new System.Drawing.Size(44, 23);
            this.btnRefDirPath.TabIndex = 2;
            this.btnRefDirPath.Text = "ref";
            this.btnRefDirPath.UseVisualStyleBackColor = true;
            // 
            // txtOutput
            // 
            this.txtOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtOutput.Location = new System.Drawing.Point(12, 205);
            this.txtOutput.MaxLength = 0;
            this.txtOutput.Multiline = true;
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtOutput.Size = new System.Drawing.Size(474, 269);
            this.txtOutput.TabIndex = 3;
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(22, 176);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(44, 23);
            this.btnRun.TabIndex = 4;
            this.btnRun.Text = "実行";
            this.btnRun.UseVisualStyleBackColor = true;
            // 
            // btnExtract
            // 
            this.btnExtract.Location = new System.Drawing.Point(429, 176);
            this.btnExtract.Name = "btnExtract";
            this.btnExtract.Size = new System.Drawing.Size(44, 23);
            this.btnExtract.TabIndex = 5;
            this.btnExtract.Text = "復元";
            this.btnExtract.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(64, 12);
            this.label2.TabIndex = 7;
            this.label2.Text = "除外フォルダ";
            // 
            // txtIgnoreDirs
            // 
            this.txtIgnoreDirs.Location = new System.Drawing.Point(22, 68);
            this.txtIgnoreDirs.Name = "txtIgnoreDirs";
            this.txtIgnoreDirs.Size = new System.Drawing.Size(471, 19);
            this.txtIgnoreDirs.TabIndex = 6;
            this.toolTip1.SetToolTip(this.txtIgnoreDirs, "複数指定時は\"/\"で区切る");
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(20, 94);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(63, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "除外ファイル";
            // 
            // txtIgnoreFiles
            // 
            this.txtIgnoreFiles.Location = new System.Drawing.Point(22, 109);
            this.txtIgnoreFiles.Name = "txtIgnoreFiles";
            this.txtIgnoreFiles.Size = new System.Drawing.Size(471, 19);
            this.txtIgnoreFiles.TabIndex = 8;
            this.toolTip1.SetToolTip(this.txtIgnoreFiles, "複数指定時は\"/\"で区切る");
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(522, 486);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtIgnoreFiles);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtIgnoreDirs);
            this.Controls.Add(this.btnExtract);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.btnRefDirPath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtDirPath);
            this.Name = "MainForm";
            this.Text = "Directory To Text";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtDirPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnRefDirPath;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Button btnExtract;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtIgnoreDirs;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtIgnoreFiles;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}

