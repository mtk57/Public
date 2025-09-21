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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnAllOpen = new System.Windows.Forms.Button();
            this.btnAllClear = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(21, 42);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(621, 199);
            this.dataGridView1.TabIndex = 0;
            // 
            // btnAllOpen
            // 
            this.btnAllOpen.Location = new System.Drawing.Point(21, 12);
            this.btnAllOpen.Name = "btnAllOpen";
            this.btnAllOpen.Size = new System.Drawing.Size(75, 23);
            this.btnAllOpen.TabIndex = 1;
            this.btnAllOpen.Text = "全て開く";
            this.btnAllOpen.UseVisualStyleBackColor = true;
            // 
            // btnAllClear
            // 
            this.btnAllClear.Location = new System.Drawing.Point(273, 12);
            this.btnAllClear.Name = "btnAllClear";
            this.btnAllClear.Size = new System.Drawing.Size(119, 23);
            this.btnAllClear.TabIndex = 2;
            this.btnAllClear.Text = "全て履歴から削除";
            this.btnAllClear.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(102, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(121, 23);
            this.button1.TabIndex = 3;
            this.button1.Text = "選択ファイルのみ開く";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // HistoryForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(666, 266);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnAllClear);
            this.Controls.Add(this.btnAllOpen);
            this.Controls.Add(this.dataGridView1);
            this.Name = "HistoryForm";
            this.Text = "HistoryForm";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnAllOpen;
        private System.Windows.Forms.Button btnAllClear;
        private System.Windows.Forms.Button button1;
    }
}