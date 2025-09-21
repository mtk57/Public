namespace SimpleExcelBookSelector
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
            this.dataGridViewResults = new System.Windows.Forms.DataGridView();
            this.clmDir = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmFile = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmSheet = new System.Windows.Forms.DataGridViewTextBoxColumn();
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
            this.clmDir,
            this.clmFile,
            this.clmSheet});
            this.dataGridViewResults.Location = new System.Drawing.Point(12, 12);
            this.dataGridViewResults.Name = "dataGridViewResults";
            this.dataGridViewResults.ReadOnly = true;
            this.dataGridViewResults.Size = new System.Drawing.Size(743, 220);
            this.dataGridViewResults.TabIndex = 6;
            // 
            // clmDir
            // 
            this.clmDir.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmDir.HeaderText = "フォルダ";
            this.clmDir.Name = "clmDir";
            this.clmDir.ReadOnly = true;
            // 
            // clmFile
            // 
            this.clmFile.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmFile.HeaderText = "ファイル";
            this.clmFile.Name = "clmFile";
            this.clmFile.ReadOnly = true;
            // 
            // clmSheet
            // 
            this.clmSheet.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmSheet.HeaderText = "シート";
            this.clmSheet.Name = "clmSheet";
            this.clmSheet.ReadOnly = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 244);
            this.Controls.Add(this.dataGridViewResults);
            this.Name = "MainForm";
            this.Text = "Simple Excel Book Selector";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResults)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.DataGridView dataGridViewResults;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmDir;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmFile;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmSheet;
    }
}

