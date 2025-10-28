namespace SimpleExcelBookSelector
{
    partial class HistoryForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.colPinned = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCheck = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.colDirectory = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFilePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colUpdatedAt = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnAllOpen = new System.Windows.Forms.Button();
            this.btnAllClear = new System.Windows.Forms.Button();
            this.btnSelectOpen = new System.Windows.Forms.Button();
            this.btnDeleteSelectedFiles = new System.Windows.Forms.Button();
            this.btnAllCheckOnOff = new System.Windows.Forms.Button();
            this.btnPinnedSelectedFiles = new System.Windows.Forms.Button();
            this.btnUnPinnedSelectedFiles = new System.Windows.Forms.Button();
            this.btnApply = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOpenAllPin = new System.Windows.Forms.Button();
            this.chkIsOpenDir = new System.Windows.Forms.CheckBox();
            this.txtDirFilter = new System.Windows.Forms.TextBox();
            this.txtFileFilter = new System.Windows.Forms.TextBox();
            this.txtFilePathFilter = new System.Windows.Forms.TextBox();
            this.txtUpdateTimeFilter = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colPinned,
            this.colCheck,
            this.colDirectory,
            this.colFileName,
            this.colFilePath,
            this.colUpdatedAt});
            this.dataGridView1.Location = new System.Drawing.Point(21, 126);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(775, 153);
            this.dataGridView1.TabIndex = 10;
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            this.dataGridView1.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_ColumnHeaderMouseClick);
            // 
            // colPinned
            // 
            this.colPinned.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colPinned.FillWeight = 20F;
            this.colPinned.HeaderText = "ピン";
            this.colPinned.Name = "colPinned";
            this.colPinned.ReadOnly = true;
            // 
            // colCheck
            // 
            this.colCheck.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colCheck.FillWeight = 20F;
            this.colCheck.HeaderText = "";
            this.colCheck.Name = "colCheck";
            // 
            // colDirectory
            // 
            this.colDirectory.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colDirectory.HeaderText = "フォルダ";
            this.colDirectory.Name = "colDirectory";
            this.colDirectory.ReadOnly = true;
            // 
            // colFileName
            // 
            this.colFileName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colFileName.HeaderText = "ファイル";
            this.colFileName.Name = "colFileName";
            this.colFileName.ReadOnly = true;
            // 
            // colFilePath
            // 
            this.colFilePath.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colFilePath.HeaderText = "ファイルパス";
            this.colFilePath.Name = "colFilePath";
            this.colFilePath.ReadOnly = true;
            // 
            // colUpdatedAt
            // 
            this.colUpdatedAt.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colUpdatedAt.FillWeight = 50F;
            this.colUpdatedAt.HeaderText = "更新日時";
            this.colUpdatedAt.Name = "colUpdatedAt";
            this.colUpdatedAt.ReadOnly = true;
            // 
            // btnAllOpen
            // 
            this.btnAllOpen.Location = new System.Drawing.Point(21, 12);
            this.btnAllOpen.Name = "btnAllOpen";
            this.btnAllOpen.Size = new System.Drawing.Size(75, 23);
            this.btnAllOpen.TabIndex = 1;
            this.btnAllOpen.Text = "全て開く";
            this.btnAllOpen.UseVisualStyleBackColor = true;
            this.btnAllOpen.Click += new System.EventHandler(this.btnAllOpen_Click);
            // 
            // btnAllClear
            // 
            this.btnAllClear.Location = new System.Drawing.Point(493, 12);
            this.btnAllClear.Name = "btnAllClear";
            this.btnAllClear.Size = new System.Drawing.Size(119, 23);
            this.btnAllClear.TabIndex = 4;
            this.btnAllClear.Text = "全て履歴から削除";
            this.btnAllClear.UseVisualStyleBackColor = true;
            this.btnAllClear.Click += new System.EventHandler(this.btnAllClear_Click);
            // 
            // btnSelectOpen
            // 
            this.btnSelectOpen.Location = new System.Drawing.Point(133, 37);
            this.btnSelectOpen.Name = "btnSelectOpen";
            this.btnSelectOpen.Size = new System.Drawing.Size(121, 23);
            this.btnSelectOpen.TabIndex = 7;
            this.btnSelectOpen.Text = "選択ファイルのみ開く";
            this.btnSelectOpen.UseVisualStyleBackColor = true;
            this.btnSelectOpen.Click += new System.EventHandler(this.btnSelectOpen_Click);
            // 
            // btnDeleteSelectedFiles
            // 
            this.btnDeleteSelectedFiles.Location = new System.Drawing.Point(618, 12);
            this.btnDeleteSelectedFiles.Name = "btnDeleteSelectedFiles";
            this.btnDeleteSelectedFiles.Size = new System.Drawing.Size(121, 23);
            this.btnDeleteSelectedFiles.TabIndex = 5;
            this.btnDeleteSelectedFiles.Text = "選択ファイルのみ削除";
            this.btnDeleteSelectedFiles.UseVisualStyleBackColor = true;
            this.btnDeleteSelectedFiles.Click += new System.EventHandler(this.btnDeleteSelectedFiles_Click);
            // 
            // btnAllCheckOnOff
            // 
            this.btnAllCheckOnOff.Location = new System.Drawing.Point(102, 12);
            this.btnAllCheckOnOff.Name = "btnAllCheckOnOff";
            this.btnAllCheckOnOff.Size = new System.Drawing.Size(74, 23);
            this.btnAllCheckOnOff.TabIndex = 2;
            this.btnAllCheckOnOff.Text = "全てチェック";
            this.btnAllCheckOnOff.UseVisualStyleBackColor = true;
            this.btnAllCheckOnOff.Click += new System.EventHandler(this.btnAllCheckOnOff_Click);
            // 
            // btnPinnedSelectedFiles
            // 
            this.btnPinnedSelectedFiles.Location = new System.Drawing.Point(493, 37);
            this.btnPinnedSelectedFiles.Name = "btnPinnedSelectedFiles";
            this.btnPinnedSelectedFiles.Size = new System.Drawing.Size(137, 23);
            this.btnPinnedSelectedFiles.TabIndex = 8;
            this.btnPinnedSelectedFiles.Text = "選択ファイルのみピン止め";
            this.btnPinnedSelectedFiles.UseVisualStyleBackColor = true;
            this.btnPinnedSelectedFiles.Click += new System.EventHandler(this.btnPinnedSelectedFiles_Click);
            // 
            // btnUnPinnedSelectedFiles
            // 
            this.btnUnPinnedSelectedFiles.Location = new System.Drawing.Point(636, 37);
            this.btnUnPinnedSelectedFiles.Name = "btnUnPinnedSelectedFiles";
            this.btnUnPinnedSelectedFiles.Size = new System.Drawing.Size(158, 23);
            this.btnUnPinnedSelectedFiles.TabIndex = 9;
            this.btnUnPinnedSelectedFiles.Text = "選択ファイルのみピン止め解除";
            this.btnUnPinnedSelectedFiles.UseVisualStyleBackColor = true;
            this.btnUnPinnedSelectedFiles.Click += new System.EventHandler(this.btnUnPinnedSelectedFiles_Click);
            // 
            // btnApply
            // 
            this.btnApply.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnApply.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnApply.Location = new System.Drawing.Point(692, 285);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(49, 23);
            this.btnApply.TabIndex = 13;
            this.btnApply.Text = "Apply";
            this.btnApply.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(747, 285);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(49, 23);
            this.btnCancel.TabIndex = 14;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOpenAllPin
            // 
            this.btnOpenAllPin.Location = new System.Drawing.Point(24, 37);
            this.btnOpenAllPin.Name = "btnOpenAllPin";
            this.btnOpenAllPin.Size = new System.Drawing.Size(103, 23);
            this.btnOpenAllPin.TabIndex = 6;
            this.btnOpenAllPin.Text = "ピン止めを全て開く";
            this.btnOpenAllPin.UseVisualStyleBackColor = true;
            this.btnOpenAllPin.Click += new System.EventHandler(this.btnOpenAllPin_Click);
            // 
            // chkIsOpenDir
            // 
            this.chkIsOpenDir.AutoSize = true;
            this.chkIsOpenDir.Location = new System.Drawing.Point(194, 12);
            this.chkIsOpenDir.Name = "chkIsOpenDir";
            this.chkIsOpenDir.Size = new System.Drawing.Size(154, 16);
            this.chkIsOpenDir.TabIndex = 3;
            this.chkIsOpenDir.Text = "ダブルクリックでフォルダを開く";
            this.chkIsOpenDir.UseVisualStyleBackColor = true;
            // 
            // txtDirFilter
            // 
            this.txtDirFilter.Location = new System.Drawing.Point(21, 101);
            this.txtDirFilter.Name = "txtDirFilter";
            this.txtDirFilter.Size = new System.Drawing.Size(75, 19);
            this.txtDirFilter.TabIndex = 15;
            // 
            // txtFileFilter
            // 
            this.txtFileFilter.Location = new System.Drawing.Point(102, 101);
            this.txtFileFilter.Name = "txtFileFilter";
            this.txtFileFilter.Size = new System.Drawing.Size(75, 19);
            this.txtFileFilter.TabIndex = 16;
            // 
            // txtFilePathFilter
            // 
            this.txtFilePathFilter.Location = new System.Drawing.Point(183, 101);
            this.txtFilePathFilter.Name = "txtFilePathFilter";
            this.txtFilePathFilter.Size = new System.Drawing.Size(75, 19);
            this.txtFilePathFilter.TabIndex = 17;
            // 
            // txtUpdateTimeFilter
            // 
            this.txtUpdateTimeFilter.Location = new System.Drawing.Point(264, 101);
            this.txtUpdateTimeFilter.Name = "txtUpdateTimeFilter";
            this.txtUpdateTimeFilter.Size = new System.Drawing.Size(75, 19);
            this.txtUpdateTimeFilter.TabIndex = 18;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 86);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 12);
            this.label1.TabIndex = 19;
            this.label1.Text = "フォルダ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(102, 86);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(39, 12);
            this.label2.TabIndex = 20;
            this.label2.Text = "ファイル";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(181, 86);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 12);
            this.label3.TabIndex = 21;
            this.label3.Text = "ファイルパス";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(262, 86);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 22;
            this.label4.Text = "更新日時";
            // 
            // HistoryForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(808, 320);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtUpdateTimeFilter);
            this.Controls.Add(this.txtFilePathFilter);
            this.Controls.Add(this.txtFileFilter);
            this.Controls.Add(this.txtDirFilter);
            this.Controls.Add(this.chkIsOpenDir);
            this.Controls.Add(this.btnOpenAllPin);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.btnUnPinnedSelectedFiles);
            this.Controls.Add(this.btnPinnedSelectedFiles);
            this.Controls.Add(this.btnAllCheckOnOff);
            this.Controls.Add(this.btnDeleteSelectedFiles);
            this.Controls.Add(this.btnSelectOpen);
            this.Controls.Add(this.btnAllClear);
            this.Controls.Add(this.btnAllOpen);
            this.Controls.Add(this.dataGridView1);
            this.Name = "HistoryForm";
            this.Text = "履歴管理";
            this.Load += new System.EventHandler(this.HistoryForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnAllOpen;
        private System.Windows.Forms.Button btnAllClear;
        private System.Windows.Forms.Button btnSelectOpen;
        private System.Windows.Forms.Button btnDeleteSelectedFiles;
        private System.Windows.Forms.Button btnAllCheckOnOff;
        private System.Windows.Forms.Button btnPinnedSelectedFiles;
        private System.Windows.Forms.Button btnUnPinnedSelectedFiles;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOpenAllPin;
        private System.Windows.Forms.DataGridViewTextBoxColumn colPinned;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colCheck;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDirectory;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFileName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFilePath;
        private System.Windows.Forms.DataGridViewTextBoxColumn colUpdatedAt;
        private System.Windows.Forms.CheckBox chkIsOpenDir;
        private System.Windows.Forms.TextBox txtDirFilter;
        private System.Windows.Forms.TextBox txtFileFilter;
        private System.Windows.Forms.TextBox txtFilePathFilter;
        private System.Windows.Forms.TextBox txtUpdateTimeFilter;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}
