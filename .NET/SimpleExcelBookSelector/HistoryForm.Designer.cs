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
            this.btnAllOpen = new System.Windows.Forms.Button();
            this.btnAllClear = new System.Windows.Forms.Button();
            this.btnSelectOpen = new System.Windows.Forms.Button();
            this.btnDeleteSelectedFiles = new System.Windows.Forms.Button();
            this.btnAllCheckOnOff = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnClearFilter = new System.Windows.Forms.Button();
            this.btnPinnedSelectedFiles = new System.Windows.Forms.Button();
            this.btnUnPinnedSelectedFiles = new System.Windows.Forms.Button();
            this.btnApply = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOpenAllPin = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(21, 66);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(775, 127);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            this.dataGridView1.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_ColumnHeaderMouseClick);
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
            this.btnAllClear.TabIndex = 2;
            this.btnAllClear.Text = "全て履歴から削除";
            this.btnAllClear.UseVisualStyleBackColor = true;
            this.btnAllClear.Click += new System.EventHandler(this.btnAllClear_Click);
            // 
            // btnSelectOpen
            // 
            this.btnSelectOpen.Location = new System.Drawing.Point(205, 12);
            this.btnSelectOpen.Name = "btnSelectOpen";
            this.btnSelectOpen.Size = new System.Drawing.Size(121, 23);
            this.btnSelectOpen.TabIndex = 3;
            this.btnSelectOpen.Text = "選択ファイルのみ開く";
            this.btnSelectOpen.UseVisualStyleBackColor = true;
            this.btnSelectOpen.Click += new System.EventHandler(this.btnSelectOpen_Click);
            // 
            // btnDeleteSelectedFiles
            // 
            this.btnDeleteSelectedFiles.Location = new System.Drawing.Point(618, 12);
            this.btnDeleteSelectedFiles.Name = "btnDeleteSelectedFiles";
            this.btnDeleteSelectedFiles.Size = new System.Drawing.Size(121, 23);
            this.btnDeleteSelectedFiles.TabIndex = 4;
            this.btnDeleteSelectedFiles.Text = "選択ファイルのみ削除";
            this.btnDeleteSelectedFiles.UseVisualStyleBackColor = true;
            this.btnDeleteSelectedFiles.Click += new System.EventHandler(this.btnDeleteSelectedFiles_Click);
            // 
            // btnAllCheckOnOff
            // 
            this.btnAllCheckOnOff.Location = new System.Drawing.Point(332, 12);
            this.btnAllCheckOnOff.Name = "btnAllCheckOnOff";
            this.btnAllCheckOnOff.Size = new System.Drawing.Size(74, 23);
            this.btnAllCheckOnOff.TabIndex = 5;
            this.btnAllCheckOnOff.Text = "全てチェック";
            this.btnAllCheckOnOff.UseVisualStyleBackColor = true;
            this.btnAllCheckOnOff.Click += new System.EventHandler(this.btnAllCheckOnOff_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(63, 39);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(368, 19);
            this.textBox1.TabIndex = 6;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 12);
            this.label1.TabIndex = 7;
            this.label1.Text = "フィルタ";
            // 
            // btnClearFilter
            // 
            this.btnClearFilter.Location = new System.Drawing.Point(437, 37);
            this.btnClearFilter.Name = "btnClearFilter";
            this.btnClearFilter.Size = new System.Drawing.Size(48, 23);
            this.btnClearFilter.TabIndex = 8;
            this.btnClearFilter.Text = "クリア";
            this.btnClearFilter.UseVisualStyleBackColor = true;
            this.btnClearFilter.Click += new System.EventHandler(this.btnClearFilter_Click);
            // 
            // btnPinnedSelectedFiles
            // 
            this.btnPinnedSelectedFiles.Location = new System.Drawing.Point(493, 37);
            this.btnPinnedSelectedFiles.Name = "btnPinnedSelectedFiles";
            this.btnPinnedSelectedFiles.Size = new System.Drawing.Size(137, 23);
            this.btnPinnedSelectedFiles.TabIndex = 9;
            this.btnPinnedSelectedFiles.Text = "選択ファイルのみピン止め";
            this.btnPinnedSelectedFiles.UseVisualStyleBackColor = true;
            this.btnPinnedSelectedFiles.Click += new System.EventHandler(this.btnPinnedSelectedFiles_Click);
            // 
            // btnUnPinnedSelectedFiles
            // 
            this.btnUnPinnedSelectedFiles.Location = new System.Drawing.Point(636, 37);
            this.btnUnPinnedSelectedFiles.Name = "btnUnPinnedSelectedFiles";
            this.btnUnPinnedSelectedFiles.Size = new System.Drawing.Size(158, 23);
            this.btnUnPinnedSelectedFiles.TabIndex = 10;
            this.btnUnPinnedSelectedFiles.Text = "選択ファイルのみピン止め解除";
            this.btnUnPinnedSelectedFiles.UseVisualStyleBackColor = true;
            this.btnUnPinnedSelectedFiles.Click += new System.EventHandler(this.btnUnPinnedSelectedFiles_Click);
            // 
            // btnApply
            // 
            this.btnApply.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnApply.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnApply.Location = new System.Drawing.Point(692, 207);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(49, 23);
            this.btnApply.TabIndex = 11;
            this.btnApply.Text = "Apply";
            this.btnApply.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(747, 207);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(49, 23);
            this.btnCancel.TabIndex = 12;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnOpenAllPin
            // 
            this.btnOpenAllPin.Location = new System.Drawing.Point(102, 12);
            this.btnOpenAllPin.Name = "btnOpenAllPin";
            this.btnOpenAllPin.Size = new System.Drawing.Size(103, 23);
            this.btnOpenAllPin.TabIndex = 13;
            this.btnOpenAllPin.Text = "ピン止めを全て開く";
            this.btnOpenAllPin.UseVisualStyleBackColor = true;
            this.btnOpenAllPin.Click += new System.EventHandler(this.btnOpenAllPin_Click);
            // 
            // HistoryForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(808, 242);
            this.Controls.Add(this.btnOpenAllPin);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.btnUnPinnedSelectedFiles);
            this.Controls.Add(this.btnPinnedSelectedFiles);
            this.Controls.Add(this.btnClearFilter);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
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
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnClearFilter;
        private System.Windows.Forms.Button btnPinnedSelectedFiles;
        private System.Windows.Forms.Button btnUnPinnedSelectedFiles;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOpenAllPin;
    }
}