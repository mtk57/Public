namespace SimpleMethodCallListCreator.Forms
{
    partial class CollectFilesForm
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
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.pbProgress = new System.Windows.Forms.ProgressBar();
            this.lblFailed = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtStartMethod = new System.Windows.Forms.TextBox();
            this.txtStartSrcFilePath = new System.Windows.Forms.TextBox();
            this.txtMethodListPath = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnRefMethodListPath = new System.Windows.Forms.Button();
            this.btnRun = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnRefStartSrcFilePath = new System.Windows.Forms.Button();
            this.txtCollectDirPath = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnRefCollectDirPath = new System.Windows.Forms.Button();
            this.txtSrcRootDirPath = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.btnRefSrcRootDirPath = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(537, 383);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 12);
            this.lblStatus.TabIndex = 49;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(381, 341);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 33);
            this.btnCancel.TabIndex = 48;
            this.btnCancel.Text = "中止";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // pbProgress
            // 
            this.pbProgress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pbProgress.Location = new System.Drawing.Point(269, 385);
            this.pbProgress.Name = "pbProgress";
            this.pbProgress.Size = new System.Drawing.Size(266, 10);
            this.pbProgress.TabIndex = 47;
            this.pbProgress.Visible = false;
            // 
            // lblFailed
            // 
            this.lblFailed.AutoSize = true;
            this.lblFailed.ForeColor = System.Drawing.Color.Red;
            this.lblFailed.Location = new System.Drawing.Point(471, 351);
            this.lblFailed.Name = "lblFailed";
            this.lblFailed.Size = new System.Drawing.Size(88, 12);
            this.lblFailed.TabIndex = 45;
            this.lblFailed.Text = "メソッド特定失敗:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(146, 196);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(318, 12);
            this.label4.TabIndex = 39;
            this.label4.Text = "※同名メソッドがある場合はメソッドリストのシグネチャを記載すること";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(37, 196);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(74, 12);
            this.label3.TabIndex = 38;
            this.label3.Text = "開始メソッド名";
            // 
            // txtStartMethod
            // 
            this.txtStartMethod.Location = new System.Drawing.Point(39, 215);
            this.txtStartMethod.Name = "txtStartMethod";
            this.txtStartMethod.Size = new System.Drawing.Size(669, 19);
            this.txtStartMethod.TabIndex = 37;
            // 
            // txtStartSrcFilePath
            // 
            this.txtStartSrcFilePath.Location = new System.Drawing.Point(39, 151);
            this.txtStartSrcFilePath.Name = "txtStartSrcFilePath";
            this.txtStartSrcFilePath.Size = new System.Drawing.Size(669, 19);
            this.txtStartSrcFilePath.TabIndex = 36;
            // 
            // txtMethodListPath
            // 
            this.txtMethodListPath.Location = new System.Drawing.Point(39, 36);
            this.txtMethodListPath.Name = "txtMethodListPath";
            this.txtMethodListPath.Size = new System.Drawing.Size(669, 19);
            this.txtMethodListPath.TabIndex = 35;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(37, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(81, 12);
            this.label2.TabIndex = 33;
            this.label2.Text = "メソッドリストパス";
            // 
            // btnRefMethodListPath
            // 
            this.btnRefMethodListPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefMethodListPath.Location = new System.Drawing.Point(718, 36);
            this.btnRefMethodListPath.Name = "btnRefMethodListPath";
            this.btnRefMethodListPath.Size = new System.Drawing.Size(45, 23);
            this.btnRefMethodListPath.TabIndex = 34;
            this.btnRefMethodListPath.Text = "参照";
            this.btnRefMethodListPath.UseVisualStyleBackColor = true;
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(300, 341);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 33);
            this.btnRun.TabIndex = 32;
            this.btnRun.Text = "実行";
            this.btnRun.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(37, 132);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 12);
            this.label1.TabIndex = 30;
            this.label1.Text = "開始ソースファイルパス";
            // 
            // btnRefStartSrcFilePath
            // 
            this.btnRefStartSrcFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefStartSrcFilePath.Location = new System.Drawing.Point(718, 147);
            this.btnRefStartSrcFilePath.Name = "btnRefStartSrcFilePath";
            this.btnRefStartSrcFilePath.Size = new System.Drawing.Size(45, 23);
            this.btnRefStartSrcFilePath.TabIndex = 31;
            this.btnRefStartSrcFilePath.Text = "参照";
            this.btnRefStartSrcFilePath.UseVisualStyleBackColor = true;
            // 
            // txtCollectDirPath
            // 
            this.txtCollectDirPath.Location = new System.Drawing.Point(39, 278);
            this.txtCollectDirPath.Name = "txtCollectDirPath";
            this.txtCollectDirPath.Size = new System.Drawing.Size(669, 19);
            this.txtCollectDirPath.TabIndex = 52;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(37, 259);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 12);
            this.label5.TabIndex = 50;
            this.label5.Text = "収集フォルダパス";
            // 
            // btnRefCollectDirPath
            // 
            this.btnRefCollectDirPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefCollectDirPath.Location = new System.Drawing.Point(718, 274);
            this.btnRefCollectDirPath.Name = "btnRefCollectDirPath";
            this.btnRefCollectDirPath.Size = new System.Drawing.Size(45, 23);
            this.btnRefCollectDirPath.TabIndex = 51;
            this.btnRefCollectDirPath.Text = "参照";
            this.btnRefCollectDirPath.UseVisualStyleBackColor = true;
            // 
            // txtSrcRootDirPath
            // 
            this.txtSrcRootDirPath.Location = new System.Drawing.Point(39, 92);
            this.txtSrcRootDirPath.Name = "txtSrcRootDirPath";
            this.txtSrcRootDirPath.Size = new System.Drawing.Size(669, 19);
            this.txtSrcRootDirPath.TabIndex = 55;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(37, 73);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(115, 12);
            this.label6.TabIndex = 53;
            this.label6.Text = "ソースルートフォルダパス";
            // 
            // btnRefSrcRootDirPath
            // 
            this.btnRefSrcRootDirPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefSrcRootDirPath.Location = new System.Drawing.Point(718, 88);
            this.btnRefSrcRootDirPath.Name = "btnRefSrcRootDirPath";
            this.btnRefSrcRootDirPath.Size = new System.Drawing.Size(45, 23);
            this.btnRefSrcRootDirPath.TabIndex = 54;
            this.btnRefSrcRootDirPath.Text = "参照";
            this.btnRefSrcRootDirPath.UseVisualStyleBackColor = true;
            // 
            // CollectFilesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(804, 421);
            this.Controls.Add(this.txtSrcRootDirPath);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.btnRefSrcRootDirPath);
            this.Controls.Add(this.txtCollectDirPath);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnRefCollectDirPath);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.pbProgress);
            this.Controls.Add(this.lblFailed);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtStartMethod);
            this.Controls.Add(this.txtStartSrcFilePath);
            this.Controls.Add(this.txtMethodListPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnRefMethodListPath);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnRefStartSrcFilePath);
            this.Name = "CollectFilesForm";
            this.Text = "ファイル収集";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ProgressBar pbProgress;
        private System.Windows.Forms.Label lblFailed;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtStartMethod;
        private System.Windows.Forms.TextBox txtStartSrcFilePath;
        private System.Windows.Forms.TextBox txtMethodListPath;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnRefMethodListPath;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnRefStartSrcFilePath;
        private System.Windows.Forms.TextBox txtCollectDirPath;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnRefCollectDirPath;
        private System.Windows.Forms.TextBox txtSrcRootDirPath;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnRefSrcRootDirPath;
    }
}
