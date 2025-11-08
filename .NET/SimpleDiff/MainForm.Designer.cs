namespace SimpleDiff
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
      private void InitializeComponent()
        {
            this.label3 = new System.Windows.Forms.Label();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnRefDirSrc = new System.Windows.Forms.Button();
            this.txtDirPathSrc = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chkEnableSubDir = new System.Windows.Forms.CheckBox();
            this.btnRefDirDst = new System.Windows.Forms.Button();
            this.txtDirPathDst = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.clmDirPathSrc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmFileNameSrc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmDirPathDst = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmFileNameDst = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblInfo = new System.Windows.Forms.Label();
            this.btnStop = new System.Windows.Forms.Button();
            this.btnRefWinMerge = new System.Windows.Forms.Button();
            this.txtWinMergePath = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 44);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(247, 12);
            this.label3.TabIndex = 0;
            this.label3.Text = "DataGirdViewの行をダブルクリックでWinMerge起動";
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(222, 193);
            this.btnStart.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(64, 24);
            this.btnStart.TabIndex = 8;
            this.btnStart.Text = "開始";
            this.btnStart.UseVisualStyleBackColor = true;
            // 
            // btnRefDirSrc
            // 
            this.btnRefDirSrc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefDirSrc.Location = new System.Drawing.Point(824, 99);
            this.btnRefDirSrc.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnRefDirSrc.Name = "btnRefDirSrc";
            this.btnRefDirSrc.Size = new System.Drawing.Size(64, 19);
            this.btnRefDirSrc.TabIndex = 4;
            this.btnRefDirSrc.Text = "参照";
            this.btnRefDirSrc.UseVisualStyleBackColor = true;
            // 
            // txtDirPathSrc
            // 
            this.txtDirPathSrc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDirPathSrc.Location = new System.Drawing.Point(26, 99);
            this.txtDirPathSrc.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtDirPathSrc.Name = "txtDirPathSrc";
            this.txtDirPathSrc.Size = new System.Drawing.Size(792, 19);
            this.txtDirPathSrc.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(26, 85);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(95, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "比較元フォルダパス";
            // 
            // chkEnableSubDir
            // 
            this.chkEnableSubDir.AutoSize = true;
            this.chkEnableSubDir.Location = new System.Drawing.Point(28, 198);
            this.chkEnableSubDir.Name = "chkEnableSubDir";
            this.chkEnableSubDir.Size = new System.Drawing.Size(118, 16);
            this.chkEnableSubDir.TabIndex = 7;
            this.chkEnableSubDir.Text = "サブフォルダも含める";
            this.chkEnableSubDir.UseVisualStyleBackColor = true;
            // 
            // btnRefDirDst
            // 
            this.btnRefDirDst.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefDirDst.Location = new System.Drawing.Point(824, 150);
            this.btnRefDirDst.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnRefDirDst.Name = "btnRefDirDst";
            this.btnRefDirDst.Size = new System.Drawing.Size(64, 19);
            this.btnRefDirDst.TabIndex = 6;
            this.btnRefDirDst.Text = "参照";
            this.btnRefDirDst.UseVisualStyleBackColor = true;
            // 
            // txtDirPathDst
            // 
            this.txtDirPathDst.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDirPathDst.Location = new System.Drawing.Point(26, 150);
            this.txtDirPathDst.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtDirPathDst.Name = "txtDirPathDst";
            this.txtDirPathDst.Size = new System.Drawing.Size(792, 19);
            this.txtDirPathDst.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 136);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 12);
            this.label2.TabIndex = 0;
            this.label2.Text = "比較先フォルダパス";
            // 
            // dataGridView
            // 
            this.dataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.clmDirPathSrc,
            this.clmFileNameSrc,
            this.clmDirPathDst,
            this.clmFileNameDst});
            this.dataGridView.Location = new System.Drawing.Point(26, 259);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowTemplate.Height = 21;
            this.dataGridView.Size = new System.Drawing.Size(862, 253);
            this.dataGridView.TabIndex = 9;
            // 
            // clmDirPathSrc
            // 
            this.clmDirPathSrc.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmDirPathSrc.HeaderText = "比較元フォルダパス";
            this.clmDirPathSrc.Name = "clmDirPathSrc";
            // 
            // clmFileNameSrc
            // 
            this.clmFileNameSrc.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmFileNameSrc.HeaderText = "比較元ファイル名";
            this.clmFileNameSrc.Name = "clmFileNameSrc";
            // 
            // clmDirPathDst
            // 
            this.clmDirPathDst.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmDirPathDst.HeaderText = "比較先フォルダパス";
            this.clmDirPathDst.Name = "clmDirPathDst";
            // 
            // clmFileNameDst
            // 
            this.clmFileNameDst.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmFileNameDst.HeaderText = "比較先ファイル名";
            this.clmFileNameDst.Name = "clmFileNameDst";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(404, 199);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(194, 15);
            this.progressBar.TabIndex = 10;
            // 
            // lblInfo
            // 
            this.lblInfo.AutoSize = true;
            this.lblInfo.Location = new System.Drawing.Point(605, 199);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(0, 12);
            this.lblInfo.TabIndex = 11;
            // 
            // btnStop
            // 
            this.btnStop.Location = new System.Drawing.Point(292, 193);
            this.btnStop.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new System.Drawing.Size(64, 24);
            this.btnStop.TabIndex = 12;
            this.btnStop.Text = "中止";
            this.btnStop.UseVisualStyleBackColor = true;
            // 
            // btnRefWinMerge
            // 
            this.btnRefWinMerge.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefWinMerge.Location = new System.Drawing.Point(822, 23);
            this.btnRefWinMerge.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnRefWinMerge.Name = "btnRefWinMerge";
            this.btnRefWinMerge.Size = new System.Drawing.Size(64, 19);
            this.btnRefWinMerge.TabIndex = 15;
            this.btnRefWinMerge.Text = "参照";
            this.btnRefWinMerge.UseVisualStyleBackColor = true;
            // 
            // txtWinMergePath
            // 
            this.txtWinMergePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtWinMergePath.Location = new System.Drawing.Point(24, 23);
            this.txtWinMergePath.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtWinMergePath.Name = "txtWinMergePath";
            this.txtWinMergePath.Size = new System.Drawing.Size(792, 19);
            this.txtWinMergePath.TabIndex = 14;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(24, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(73, 12);
            this.label4.TabIndex = 13;
            this.label4.Text = "WinMergeパス";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(914, 539);
            this.Controls.Add(this.btnRefWinMerge);
            this.Controls.Add(this.txtWinMergePath);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnStop);
            this.Controls.Add(this.lblInfo);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnRefDirDst);
            this.Controls.Add(this.txtDirPathDst);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.chkEnableSubDir);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnRefDirSrc);
            this.Controls.Add(this.txtDirPathSrc);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView);
            this.Name = "MainForm";
            this.Text = "Simple Diff";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnRefDirSrc;
        private System.Windows.Forms.TextBox txtDirPathSrc;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkEnableSubDir;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Button btnRefDirDst;
        private System.Windows.Forms.TextBox txtDirPathDst;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmDirPathSrc;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmFileNameSrc;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmDirPathDst;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmFileNameDst;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.Button btnStop;
        private System.Windows.Forms.Button btnRefWinMerge;
        private System.Windows.Forms.TextBox txtWinMergePath;
        private System.Windows.Forms.Label label4;
    }
}