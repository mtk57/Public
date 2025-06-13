namespace SimpleExcelDiff
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
            this.lblStatus = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnProcess = new System.Windows.Forms.Button();
            this.btnBrowseSrc = new System.Windows.Forms.Button();
            this.txtPathSrc = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chkEnableSubDir = new System.Windows.Forms.CheckBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnBrowseDst = new System.Windows.Forms.Button();
            this.txtPathDst = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(20, 190);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(53, 12);
            this.lblStatus.TabIndex = 17;
            this.lblStatus.Text = "準備完了";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 222);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 15;
            this.label3.Text = "処理結果";
            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(260, 135);
            this.btnProcess.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(64, 24);
            this.btnProcess.TabIndex = 14;
            this.btnProcess.Text = "処理開始";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // btnBrowseSrc
            // 
            this.btnBrowseSrc.Location = new System.Drawing.Point(478, 36);
            this.btnBrowseSrc.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnBrowseSrc.Name = "btnBrowseSrc";
            this.btnBrowseSrc.Size = new System.Drawing.Size(64, 19);
            this.btnBrowseSrc.TabIndex = 11;
            this.btnBrowseSrc.Text = "参照";
            this.btnBrowseSrc.UseVisualStyleBackColor = true;
            this.btnBrowseSrc.Click += new System.EventHandler(this.btnBrowseSrc_Click);
            // 
            // txtPathSrc
            // 
            this.txtPathSrc.Location = new System.Drawing.Point(22, 36);
            this.txtPathSrc.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtPathSrc.Name = "txtPathSrc";
            this.txtPathSrc.Size = new System.Drawing.Size(451, 19);
            this.txtPathSrc.TabIndex = 10;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(191, 12);
            this.label1.TabIndex = 9;
            this.label1.Text = "Excelフォルダ or ファイル パス (比較元)";
            // 
            // chkEnableSubDir
            // 
            this.chkEnableSubDir.AutoSize = true;
            this.chkEnableSubDir.Location = new System.Drawing.Point(24, 135);
            this.chkEnableSubDir.Name = "chkEnableSubDir";
            this.chkEnableSubDir.Size = new System.Drawing.Size(118, 16);
            this.chkEnableSubDir.TabIndex = 20;
            this.chkEnableSubDir.Text = "サブフォルダも含める";
            this.chkEnableSubDir.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(24, 237);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(528, 216);
            this.dataGridView1.TabIndex = 21;
            // 
            // btnBrowseDst
            // 
            this.btnBrowseDst.Location = new System.Drawing.Point(478, 87);
            this.btnBrowseDst.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnBrowseDst.Name = "btnBrowseDst";
            this.btnBrowseDst.Size = new System.Drawing.Size(64, 19);
            this.btnBrowseDst.TabIndex = 24;
            this.btnBrowseDst.Text = "参照";
            this.btnBrowseDst.UseVisualStyleBackColor = true;
            // 
            // txtPathDst
            // 
            this.txtPathDst.Location = new System.Drawing.Point(22, 87);
            this.txtPathDst.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtPathDst.Name = "txtPathDst";
            this.txtPathDst.Size = new System.Drawing.Size(451, 19);
            this.txtPathDst.TabIndex = 23;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(191, 12);
            this.label2.TabIndex = 22;
            this.label2.Text = "Excelフォルダ or ファイル パス (比較先)";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(576, 474);
            this.Controls.Add(this.btnBrowseDst);
            this.Controls.Add(this.txtPathDst);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.chkEnableSubDir);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.btnBrowseSrc);
            this.Controls.Add(this.txtPathSrc);
            this.Controls.Add(this.label1);
            this.Name = "MainForm";
            this.Text = "Simple Excel Diff";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.Button btnBrowseSrc;
        private System.Windows.Forms.TextBox txtPathSrc;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkEnableSubDir;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnBrowseDst;
        private System.Windows.Forms.TextBox txtPathDst;
        private System.Windows.Forms.Label label2;
    }
}