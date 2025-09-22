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
            this.chkEnableSheetSelectMode = new System.Windows.Forms.CheckBox();
            this.chkEnableAutoUpdateMode = new System.Windows.Forms.CheckBox();
            this.textAutoUpdateSec = new System.Windows.Forms.TextBox();
            this.cmbHistory = new System.Windows.Forms.ComboBox();
            this.btnForceUpdate = new System.Windows.Forms.Button();
            this.btnHistory = new System.Windows.Forms.Button();
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
            this.dataGridViewResults.Location = new System.Drawing.Point(12, 65);
            this.dataGridViewResults.Name = "dataGridViewResults";
            this.dataGridViewResults.ReadOnly = true;
            this.dataGridViewResults.Size = new System.Drawing.Size(693, 151);
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
            // chkEnableSheetSelectMode
            // 
            this.chkEnableSheetSelectMode.AutoSize = true;
            this.chkEnableSheetSelectMode.Location = new System.Drawing.Point(12, 12);
            this.chkEnableSheetSelectMode.Name = "chkEnableSheetSelectMode";
            this.chkEnableSheetSelectMode.Size = new System.Drawing.Size(104, 16);
            this.chkEnableSheetSelectMode.TabIndex = 7;
            this.chkEnableSheetSelectMode.Text = "シート選択モード";
            this.chkEnableSheetSelectMode.UseVisualStyleBackColor = true;
            // 
            // chkEnableAutoUpdateMode
            // 
            this.chkEnableAutoUpdateMode.AutoSize = true;
            this.chkEnableAutoUpdateMode.Location = new System.Drawing.Point(215, 11);
            this.chkEnableAutoUpdateMode.Name = "chkEnableAutoUpdateMode";
            this.chkEnableAutoUpdateMode.Size = new System.Drawing.Size(196, 16);
            this.chkEnableAutoUpdateMode.TabIndex = 8;
            this.chkEnableAutoUpdateMode.Text = "自動更新間隔(秒単位。最小値=1)";
            this.chkEnableAutoUpdateMode.UseVisualStyleBackColor = true;
            // 
            // textAutoUpdateSec
            // 
            this.textAutoUpdateSec.Location = new System.Drawing.Point(159, 9);
            this.textAutoUpdateSec.Name = "textAutoUpdateSec";
            this.textAutoUpdateSec.Size = new System.Drawing.Size(50, 19);
            this.textAutoUpdateSec.TabIndex = 9;
            this.textAutoUpdateSec.Text = "1";
            // 
            // cmbHistory
            // 
            this.cmbHistory.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbHistory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbHistory.FormattingEnabled = true;
            this.cmbHistory.Location = new System.Drawing.Point(12, 39);
            this.cmbHistory.Name = "cmbHistory";
            this.cmbHistory.Size = new System.Drawing.Size(693, 20);
            this.cmbHistory.TabIndex = 10;
            // 
            // btnForceUpdate
            // 
            this.btnForceUpdate.Location = new System.Drawing.Point(428, 7);
            this.btnForceUpdate.Name = "btnForceUpdate";
            this.btnForceUpdate.Size = new System.Drawing.Size(75, 23);
            this.btnForceUpdate.TabIndex = 11;
            this.btnForceUpdate.Text = "強制更新";
            this.btnForceUpdate.UseVisualStyleBackColor = true;
            // 
            // btnHistory
            // 
            this.btnHistory.Location = new System.Drawing.Point(634, 4);
            this.btnHistory.Name = "btnHistory";
            this.btnHistory.Size = new System.Drawing.Size(71, 23);
            this.btnHistory.TabIndex = 12;
            this.btnHistory.Text = "履歴管理";
            this.btnHistory.UseVisualStyleBackColor = true;
            this.btnHistory.Click += new System.EventHandler(this.btnHistory_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(717, 228);
            this.Controls.Add(this.btnHistory);
            this.Controls.Add(this.btnForceUpdate);
            this.Controls.Add(this.cmbHistory);
            this.Controls.Add(this.textAutoUpdateSec);
            this.Controls.Add(this.chkEnableAutoUpdateMode);
            this.Controls.Add(this.chkEnableSheetSelectMode);
            this.Controls.Add(this.dataGridViewResults);
            this.Name = "MainForm";
            this.Text = "Simple Excel Book Selector";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResults)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.DataGridView dataGridViewResults;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmDir;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmFile;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmSheet;
        private System.Windows.Forms.CheckBox chkEnableSheetSelectMode;
        private System.Windows.Forms.CheckBox chkEnableAutoUpdateMode;
        private System.Windows.Forms.TextBox textAutoUpdateSec;
        private System.Windows.Forms.ComboBox cmbHistory;
        private System.Windows.Forms.Button btnForceUpdate;
        private System.Windows.Forms.Button btnHistory;
    }
}

