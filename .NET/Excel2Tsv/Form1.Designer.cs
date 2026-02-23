namespace Excel2Tsv
{
    partial class Form1
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
            this.txtExcelFilePath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnRefExcelFilePath = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.txtTsvDirPath = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.btnRefTsvDirPath = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.btnStart = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // txtExcelFilePath
            // 
            this.txtExcelFilePath.Location = new System.Drawing.Point(29, 46);
            this.txtExcelFilePath.Name = "txtExcelFilePath";
            this.txtExcelFilePath.Size = new System.Drawing.Size(624, 19);
            this.txtExcelFilePath.TabIndex = 0;
            this.toolTip1.SetToolTip(this.txtExcelFilePath, "ドラッグアンドドロップ可");
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(187, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "Excel ファイルパス (xlsはNG)　※必須";
            // 
            // btnRefExcelFilePath
            // 
            this.btnRefExcelFilePath.Location = new System.Drawing.Point(659, 46);
            this.btnRefExcelFilePath.Name = "btnRefExcelFilePath";
            this.btnRefExcelFilePath.Size = new System.Drawing.Size(44, 23);
            this.btnRefExcelFilePath.TabIndex = 2;
            this.btnRefExcelFilePath.Text = "参照";
            this.btnRefExcelFilePath.UseVisualStyleBackColor = true;
            // 
            // txtTsvDirPath
            // 
            this.txtTsvDirPath.Location = new System.Drawing.Point(29, 94);
            this.txtTsvDirPath.Name = "txtTsvDirPath";
            this.txtTsvDirPath.Size = new System.Drawing.Size(624, 19);
            this.txtTsvDirPath.TabIndex = 5;
            this.toolTip1.SetToolTip(this.txtTsvDirPath, "ドラッグアンドドロップ可");
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(29, 175);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(674, 252);
            this.dataGridView1.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 160);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(151, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "シート名とTSVファイルの紐づけ";
            // 
            // btnRefTsvDirPath
            // 
            this.btnRefTsvDirPath.Location = new System.Drawing.Point(659, 94);
            this.btnRefTsvDirPath.Name = "btnRefTsvDirPath";
            this.btnRefTsvDirPath.Size = new System.Drawing.Size(44, 23);
            this.btnRefTsvDirPath.TabIndex = 7;
            this.btnRefTsvDirPath.Text = "参照";
            this.btnRefTsvDirPath.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(27, 79);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(153, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "TSV 出力フォルダパス　※任意";
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(278, 131);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(176, 23);
            this.btnStart.TabIndex = 8;
            this.btnStart.Text = "開始";
            this.btnStart.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(730, 450);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnRefTsvDirPath);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtTsvDirPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnRefExcelFilePath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtExcelFilePath);
            this.Name = "Form1";
            this.Text = "Excel To Tsv";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtExcelFilePath;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnRefExcelFilePath;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnRefTsvDirPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtTsvDirPath;
        private System.Windows.Forms.Button btnStart;
    }
}

