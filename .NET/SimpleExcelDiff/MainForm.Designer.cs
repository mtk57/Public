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
      private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.lblStatus = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnProcess = new System.Windows.Forms.Button();
            this.btnBrowseSrc = new System.Windows.Forms.Button();
            this.txtPathSrc = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chkEnableSubDir = new System.Windows.Forms.CheckBox();
            this.btnBrowseDst = new System.Windows.Forms.Button();
            this.txtPathDst = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.chkDrawCell = new System.Windows.Forms.CheckBox();
            this.btnColorSelector = new System.Windows.Forms.Button();
            this.picCellColor = new System.Windows.Forms.PictureBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picCellColor)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblStatus
            // 
            this.lblStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(20, 521);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(53, 12);
            this.lblStatus.TabIndex = 0;
            this.lblStatus.Text = "準備完了";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 181);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 0;
            this.label3.Text = "処理結果";
            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(260, 135);
            this.btnProcess.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(64, 24);
            this.btnProcess.TabIndex = 8;
            this.btnProcess.Text = "処理開始";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // btnBrowseSrc
            // 
            this.btnBrowseSrc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowseSrc.Location = new System.Drawing.Point(901, 36);
            this.btnBrowseSrc.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnBrowseSrc.Name = "btnBrowseSrc";
            this.btnBrowseSrc.Size = new System.Drawing.Size(64, 19);
            this.btnBrowseSrc.TabIndex = 4;
            this.btnBrowseSrc.Text = "参照";
            this.btnBrowseSrc.UseVisualStyleBackColor = true;
            this.btnBrowseSrc.Click += new System.EventHandler(this.btnBrowseSrc_Click);
            // 
            // txtPathSrc
            // 
            this.txtPathSrc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPathSrc.Location = new System.Drawing.Point(22, 36);
            this.txtPathSrc.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtPathSrc.Name = "txtPathSrc";
            this.txtPathSrc.Size = new System.Drawing.Size(873, 19);
            this.txtPathSrc.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(191, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Excelフォルダ or ファイル パス (比較元)";
            // 
            // chkEnableSubDir
            // 
            this.chkEnableSubDir.AutoSize = true;
            this.chkEnableSubDir.Location = new System.Drawing.Point(24, 135);
            this.chkEnableSubDir.Name = "chkEnableSubDir";
            this.chkEnableSubDir.Size = new System.Drawing.Size(118, 16);
            this.chkEnableSubDir.TabIndex = 7;
            this.chkEnableSubDir.Text = "サブフォルダも含める";
            this.chkEnableSubDir.UseVisualStyleBackColor = true;
            // 
            // btnBrowseDst
            // 
            this.btnBrowseDst.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowseDst.Location = new System.Drawing.Point(901, 87);
            this.btnBrowseDst.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnBrowseDst.Name = "btnBrowseDst";
            this.btnBrowseDst.Size = new System.Drawing.Size(64, 19);
            this.btnBrowseDst.TabIndex = 6;
            this.btnBrowseDst.Text = "参照";
            this.btnBrowseDst.UseVisualStyleBackColor = true;
            this.btnBrowseDst.Click += new System.EventHandler(this.btnBrowseDst_Click);
            // 
            // txtPathDst
            // 
            this.txtPathDst.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPathDst.Location = new System.Drawing.Point(22, 87);
            this.txtPathDst.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtPathDst.Name = "txtPathDst";
            this.txtPathDst.Size = new System.Drawing.Size(873, 19);
            this.txtPathDst.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(191, 12);
            this.label2.TabIndex = 0;
            this.label2.Text = "Excelフォルダ or ファイル パス (比較先)";
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(22, 196);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(943, 312);
            this.dataGridView1.TabIndex = 9;
            // 
            // chkDrawCell
            // 
            this.chkDrawCell.AutoSize = true;
            this.chkDrawCell.Location = new System.Drawing.Point(17, 27);
            this.chkDrawCell.Name = "chkDrawCell";
            this.chkDrawCell.Size = new System.Drawing.Size(48, 16);
            this.chkDrawCell.TabIndex = 1;
            this.chkDrawCell.Text = "有効";
            this.chkDrawCell.UseVisualStyleBackColor = true;
            this.chkDrawCell.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // btnColorSelector
            // 
            this.btnColorSelector.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnColorSelector.Location = new System.Drawing.Point(103, 23);
            this.btnColorSelector.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnColorSelector.Name = "btnColorSelector";
            this.btnColorSelector.Size = new System.Drawing.Size(52, 19);
            this.btnColorSelector.TabIndex = 3;
            this.btnColorSelector.Text = "色選択";
            this.btnColorSelector.UseVisualStyleBackColor = true;
            this.btnColorSelector.Click += new System.EventHandler(this.btnColorSelector_Click);
            // 
            // picCellColor
            // 
            this.picCellColor.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.picCellColor.Location = new System.Drawing.Point(71, 23);
            this.picCellColor.Name = "picCellColor";
            this.picCellColor.Size = new System.Drawing.Size(23, 20);
            this.picCellColor.TabIndex = 0;
            this.picCellColor.TabStop = false;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkDrawCell);
            this.groupBox1.Controls.Add(this.btnColorSelector);
            this.groupBox1.Controls.Add(this.picCellColor);
            this.groupBox1.Location = new System.Drawing.Point(373, 113);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(161, 58);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "セル色変更";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(995, 552);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnBrowseDst);
            this.Controls.Add(this.txtPathDst);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.chkEnableSubDir);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.btnBrowseSrc);
            this.Controls.Add(this.txtPathSrc);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "MainForm";
            this.Text = "Simple Excel Diff";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picCellColor)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
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
        private System.Windows.Forms.CheckBox chkDrawCell;
        private System.Windows.Forms.Button btnColorSelector;
        private System.Windows.Forms.PictureBox picCellColor;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}