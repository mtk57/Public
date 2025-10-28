﻿namespace SimpleExcelBookSelector
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
            this.chkEnableSheetSelectMode = new System.Windows.Forms.CheckBox();
            this.chkEnableAutoUpdateMode = new System.Windows.Forms.CheckBox();
            this.textAutoUpdateSec = new System.Windows.Forms.TextBox();
            this.btnForceUpdate = new System.Windows.Forms.Button();
            this.btnHistory = new System.Windows.Forms.Button();
            this.chkIsOpenDir = new System.Windows.Forms.CheckBox();
            this.clmPinned = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmDir = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmFile = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmSheet = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmUpdated = new System.Windows.Forms.DataGridViewTextBoxColumn();
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
            this.clmPinned,
            this.clmDir,
            this.clmFile,
            this.clmSheet,
            this.clmUpdated});
            this.dataGridViewResults.Location = new System.Drawing.Point(12, 64);
            this.dataGridViewResults.MultiSelect = false;
            this.dataGridViewResults.Name = "dataGridViewResults";
            this.dataGridViewResults.ReadOnly = true;
            this.dataGridViewResults.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dataGridViewResults.Size = new System.Drawing.Size(693, 191);
            this.dataGridViewResults.TabIndex = 7;
            // 
            // chkEnableSheetSelectMode
            // 
            this.chkEnableSheetSelectMode.AutoSize = true;
            this.chkEnableSheetSelectMode.Location = new System.Drawing.Point(12, 12);
            this.chkEnableSheetSelectMode.Name = "chkEnableSheetSelectMode";
            this.chkEnableSheetSelectMode.Size = new System.Drawing.Size(104, 16);
            this.chkEnableSheetSelectMode.TabIndex = 1;
            this.chkEnableSheetSelectMode.Text = "シート選択モード";
            this.chkEnableSheetSelectMode.UseVisualStyleBackColor = true;
            // 
            // chkEnableAutoUpdateMode
            // 
            this.chkEnableAutoUpdateMode.AutoSize = true;
            this.chkEnableAutoUpdateMode.Location = new System.Drawing.Point(215, 11);
            this.chkEnableAutoUpdateMode.Name = "chkEnableAutoUpdateMode";
            this.chkEnableAutoUpdateMode.Size = new System.Drawing.Size(196, 16);
            this.chkEnableAutoUpdateMode.TabIndex = 3;
            this.chkEnableAutoUpdateMode.Text = "自動更新間隔(秒単位。最小値=1)";
            this.chkEnableAutoUpdateMode.UseVisualStyleBackColor = true;
            // 
            // textAutoUpdateSec
            // 
            this.textAutoUpdateSec.Location = new System.Drawing.Point(159, 9);
            this.textAutoUpdateSec.Name = "textAutoUpdateSec";
            this.textAutoUpdateSec.Size = new System.Drawing.Size(50, 19);
            this.textAutoUpdateSec.TabIndex = 2;
            this.textAutoUpdateSec.Text = "1";
            // 
            // btnForceUpdate
            // 
            this.btnForceUpdate.Location = new System.Drawing.Point(428, 7);
            this.btnForceUpdate.Name = "btnForceUpdate";
            this.btnForceUpdate.Size = new System.Drawing.Size(75, 23);
            this.btnForceUpdate.TabIndex = 4;
            this.btnForceUpdate.Text = "強制更新";
            this.btnForceUpdate.UseVisualStyleBackColor = true;
            // 
            // btnHistory
            // 
            this.btnHistory.Location = new System.Drawing.Point(634, 4);
            this.btnHistory.Name = "btnHistory";
            this.btnHistory.Size = new System.Drawing.Size(71, 23);
            this.btnHistory.TabIndex = 5;
            this.btnHistory.Text = "履歴管理";
            this.btnHistory.UseVisualStyleBackColor = true;
            this.btnHistory.Click += new System.EventHandler(this.btnHistory_Click);
            // 
            // chkIsOpenDir
            // 
            this.chkIsOpenDir.AutoSize = true;
            this.chkIsOpenDir.Location = new System.Drawing.Point(12, 34);
            this.chkIsOpenDir.Name = "chkIsOpenDir";
            this.chkIsOpenDir.Size = new System.Drawing.Size(154, 16);
            this.chkIsOpenDir.TabIndex = 6;
            this.chkIsOpenDir.Text = "ダブルクリックでフォルダを開く";
            this.chkIsOpenDir.UseVisualStyleBackColor = true;
            // 
            // clmPinned
            // 
            this.clmPinned.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmPinned.FillWeight = 20F;
            this.clmPinned.HeaderText = "ピン留め";
            this.clmPinned.Name = "clmPinned";
            this.clmPinned.ReadOnly = true;
            this.clmPinned.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic;
            // 
            // clmDir
            // 
            this.clmDir.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmDir.HeaderText = "フォルダ";
            this.clmDir.Name = "clmDir";
            this.clmDir.ReadOnly = true;
            this.clmDir.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic;
            // 
            // clmFile
            // 
            this.clmFile.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmFile.HeaderText = "ファイル";
            this.clmFile.Name = "clmFile";
            this.clmFile.ReadOnly = true;
            this.clmFile.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic;
            // 
            // clmSheet
            // 
            this.clmSheet.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmSheet.HeaderText = "シート";
            this.clmSheet.Name = "clmSheet";
            this.clmSheet.ReadOnly = true;
            this.clmSheet.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic;
            // 
            // clmUpdated
            // 
            this.clmUpdated.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmUpdated.FillWeight = 50F;
            this.clmUpdated.HeaderText = "更新日時";
            this.clmUpdated.Name = "clmUpdated";
            this.clmUpdated.ReadOnly = true;
            this.clmUpdated.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(717, 267);
            this.Controls.Add(this.chkIsOpenDir);
            this.Controls.Add(this.btnHistory);
            this.Controls.Add(this.btnForceUpdate);
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
        private System.Windows.Forms.CheckBox chkEnableSheetSelectMode;
        private System.Windows.Forms.CheckBox chkEnableAutoUpdateMode;
        private System.Windows.Forms.TextBox textAutoUpdateSec;
        private System.Windows.Forms.Button btnForceUpdate;
        private System.Windows.Forms.Button btnHistory;
        private System.Windows.Forms.CheckBox chkIsOpenDir;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmPinned;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmDir;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmFile;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmSheet;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmUpdated;
    }
}

