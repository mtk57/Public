namespace SimpleExcelGrep.Forms
{
    partial class OtherForm
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
            this.chkAllShape = new System.Windows.Forms.CheckBox();
            this.chkAllFormula = new System.Windows.Forms.CheckBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblFolder = new System.Windows.Forms.Label();
            this.txtFolderPath = new System.Windows.Forms.TextBox();
            this.chkSubFolders = new System.Windows.Forms.CheckBox();
            this.lblIgnoreFileSize = new System.Windows.Forms.Label();
            this.txtIgnoreFileSizeMB = new System.Windows.Forms.TextBox();
            this.lblIgnoreUnit = new System.Windows.Forms.Label();
            this.chkInvisibleSheets = new System.Windows.Forms.CheckBox();
            this.lblParallelism = new System.Windows.Forms.Label();
            this.nudParallelism = new System.Windows.Forms.NumericUpDown();
            this.chkEnableLog = new System.Windows.Forms.CheckBox();
            this.lblProgress = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.nudParallelism)).BeginInit();
            this.SuspendLayout();
            // 
            // chkAllShape
            // 
            this.chkAllShape.AutoSize = true;
            this.chkAllShape.Location = new System.Drawing.Point(24, 109);
            this.chkAllShape.Name = "chkAllShape";
            this.chkAllShape.Size = new System.Drawing.Size(67, 16);
            this.chkAllShape.TabIndex = 7;
            this.chkAllShape.Text = "全ての図";
            this.chkAllShape.UseVisualStyleBackColor = true;
            // 
            // chkAllFormula
            // 
            this.chkAllFormula.AutoSize = true;
            this.chkAllFormula.Location = new System.Drawing.Point(117, 109);
            this.chkAllFormula.Name = "chkAllFormula";
            this.chkAllFormula.Size = new System.Drawing.Size(79, 16);
            this.chkAllFormula.TabIndex = 8;
            this.chkAllFormula.Text = "全ての数式";
            this.chkAllFormula.UseVisualStyleBackColor = true;
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(24, 214);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(120, 26);
            this.btnRun.TabIndex = 13;
            this.btnRun.Text = "実行";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.BtnRun_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(162, 214);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(120, 26);
            this.btnCancel.TabIndex = 14;
            this.btnCancel.Text = "中止";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(24, 182);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(358, 12);
            this.progressBar1.TabIndex = 12;
            // 
            // lblFolder
            // 
            this.lblFolder.Location = new System.Drawing.Point(22, 16);
            this.lblFolder.Name = "lblFolder";
            this.lblFolder.Size = new System.Drawing.Size(76, 18);
            this.lblFolder.TabIndex = 0;
            this.lblFolder.Text = "フォルダパス:";
            this.lblFolder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtFolderPath
            // 
            this.txtFolderPath.Location = new System.Drawing.Point(104, 14);
            this.txtFolderPath.Name = "txtFolderPath";
            this.txtFolderPath.Size = new System.Drawing.Size(278, 19);
            this.txtFolderPath.TabIndex = 1;
            // 
            // chkSubFolders
            // 
            this.chkSubFolders.AutoSize = true;
            this.chkSubFolders.Location = new System.Drawing.Point(24, 47);
            this.chkSubFolders.Name = "chkSubFolders";
            this.chkSubFolders.Size = new System.Drawing.Size(111, 16);
            this.chkSubFolders.TabIndex = 2;
            this.chkSubFolders.Text = "サブフォルダも対象";
            this.chkSubFolders.UseVisualStyleBackColor = true;
            // 
            // lblIgnoreFileSize
            // 
            this.lblIgnoreFileSize.Location = new System.Drawing.Point(22, 76);
            this.lblIgnoreFileSize.Name = "lblIgnoreFileSize";
            this.lblIgnoreFileSize.Size = new System.Drawing.Size(116, 18);
            this.lblIgnoreFileSize.TabIndex = 0;
            this.lblIgnoreFileSize.Text = "無視ファイルサイズ:";
            this.lblIgnoreFileSize.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtIgnoreFileSizeMB
            // 
            this.txtIgnoreFileSizeMB.Location = new System.Drawing.Point(144, 74);
            this.txtIgnoreFileSizeMB.Name = "txtIgnoreFileSizeMB";
            this.txtIgnoreFileSizeMB.Size = new System.Drawing.Size(60, 19);
            this.txtIgnoreFileSizeMB.TabIndex = 4;
            this.txtIgnoreFileSizeMB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // lblIgnoreUnit
            // 
            this.lblIgnoreUnit.AutoSize = true;
            this.lblIgnoreUnit.Location = new System.Drawing.Point(207, 78);
            this.lblIgnoreUnit.Name = "lblIgnoreUnit";
            this.lblIgnoreUnit.Size = new System.Drawing.Size(30, 12);
            this.lblIgnoreUnit.TabIndex = 0;
            this.lblIgnoreUnit.Text = "(MB)";
            // 
            // chkInvisibleSheets
            // 
            this.chkInvisibleSheets.AutoSize = true;
            this.chkInvisibleSheets.Location = new System.Drawing.Point(152, 47);
            this.chkInvisibleSheets.Name = "chkInvisibleSheets";
            this.chkInvisibleSheets.Size = new System.Drawing.Size(121, 16);
            this.chkInvisibleSheets.TabIndex = 3;
            this.chkInvisibleSheets.Text = "非表示シートも対象";
            this.chkInvisibleSheets.UseVisualStyleBackColor = true;
            // 
            // lblParallelism
            // 
            this.lblParallelism.AutoSize = true;
            this.lblParallelism.Location = new System.Drawing.Point(243, 78);
            this.lblParallelism.Name = "lblParallelism";
            this.lblParallelism.Size = new System.Drawing.Size(45, 12);
            this.lblParallelism.TabIndex = 0;
            this.lblParallelism.Text = "並列数 :";
            // 
            // nudParallelism
            // 
            this.nudParallelism.Location = new System.Drawing.Point(294, 74);
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
            this.nudParallelism.Size = new System.Drawing.Size(60, 19);
            this.nudParallelism.TabIndex = 5;
            this.nudParallelism.Value = new decimal(new int[] {
            4,
            0,
            0,
            0});
            // 
            // chkEnableLog
            // 
            this.chkEnableLog.AutoSize = true;
            this.chkEnableLog.Location = new System.Drawing.Point(292, 47);
            this.chkEnableLog.Name = "chkEnableLog";
            this.chkEnableLog.Size = new System.Drawing.Size(75, 16);
            this.chkEnableLog.TabIndex = 6;
            this.chkEnableLog.Text = "ログを出力";
            this.chkEnableLog.UseVisualStyleBackColor = true;
            // 
            // lblProgress
            // 
            this.lblProgress.Location = new System.Drawing.Point(22, 161);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(360, 18);
            this.lblProgress.TabIndex = 0;
            this.lblProgress.Text = "準備完了";
            this.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // OtherForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(404, 256);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.chkEnableLog);
            this.Controls.Add(this.nudParallelism);
            this.Controls.Add(this.lblParallelism);
            this.Controls.Add(this.chkInvisibleSheets);
            this.Controls.Add(this.lblIgnoreUnit);
            this.Controls.Add(this.txtIgnoreFileSizeMB);
            this.Controls.Add(this.lblIgnoreFileSize);
            this.Controls.Add(this.chkSubFolders);
            this.Controls.Add(this.txtFolderPath);
            this.Controls.Add(this.lblFolder);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.chkAllFormula);
            this.Controls.Add(this.chkAllShape);
            this.MaximumSize = new System.Drawing.Size(420, 295);
            this.MinimumSize = new System.Drawing.Size(420, 295);
            this.Name = "OtherForm";
            this.Text = "その他";
            ((System.ComponentModel.ISupportInitialize)(this.nudParallelism)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkAllShape;
        private System.Windows.Forms.CheckBox chkAllFormula;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblFolder;
        private System.Windows.Forms.TextBox txtFolderPath;
        private System.Windows.Forms.CheckBox chkSubFolders;
        private System.Windows.Forms.Label lblIgnoreFileSize;
        private System.Windows.Forms.TextBox txtIgnoreFileSizeMB;
        private System.Windows.Forms.Label lblIgnoreUnit;
        private System.Windows.Forms.CheckBox chkInvisibleSheets;
        private System.Windows.Forms.Label lblParallelism;
        private System.Windows.Forms.NumericUpDown nudParallelism;
        private System.Windows.Forms.CheckBox chkEnableLog;
        private System.Windows.Forms.Label lblProgress;
    }
}
