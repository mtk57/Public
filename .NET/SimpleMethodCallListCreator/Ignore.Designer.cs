namespace SimpleMethodCallListCreator
{
    partial class IgnoreForm
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
            this.btnInsertRow = new System.Windows.Forms.Button();
            this.btnRemoveRow = new System.Windows.Forms.Button();
            this.clmKeyword = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmMode = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.clmRegEx = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.clmCase = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.clmKeyword,
            this.clmMode,
            this.clmRegEx,
            this.clmCase});
            this.dataGridView1.Location = new System.Drawing.Point(29, 49);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(595, 192);
            this.dataGridView1.TabIndex = 0;
            // 
            // btnInsertRow
            // 
            this.btnInsertRow.Location = new System.Drawing.Point(29, 20);
            this.btnInsertRow.Name = "btnInsertRow";
            this.btnInsertRow.Size = new System.Drawing.Size(75, 23);
            this.btnInsertRow.TabIndex = 1;
            this.btnInsertRow.Text = "行追加";
            this.btnInsertRow.UseVisualStyleBackColor = true;
            // 
            // btnRemoveRow
            // 
            this.btnRemoveRow.Location = new System.Drawing.Point(110, 20);
            this.btnRemoveRow.Name = "btnRemoveRow";
            this.btnRemoveRow.Size = new System.Drawing.Size(75, 23);
            this.btnRemoveRow.TabIndex = 2;
            this.btnRemoveRow.Text = "行削除";
            this.btnRemoveRow.UseVisualStyleBackColor = true;
            // 
            // clmKeyword
            // 
            this.clmKeyword.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmKeyword.HeaderText = "キーワード";
            this.clmKeyword.Name = "clmKeyword";
            // 
            // clmMode
            // 
            this.clmMode.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmMode.HeaderText = "モード";
            this.clmMode.Items.AddRange(new object[] {
            "始まる",
            "終わる",
            "含む",
            "一致"});
            this.clmMode.Name = "clmMode";
            // 
            // clmRegEx
            // 
            this.clmRegEx.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmRegEx.HeaderText = "正規表現";
            this.clmRegEx.Name = "clmRegEx";
            // 
            // clmCase
            // 
            this.clmCase.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmCase.HeaderText = "大小区別";
            this.clmCase.Name = "clmCase";
            // 
            // IgnoreForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(639, 263);
            this.Controls.Add(this.btnRemoveRow);
            this.Controls.Add(this.btnInsertRow);
            this.Controls.Add(this.dataGridView1);
            this.Name = "IgnoreForm";
            this.Text = "除外指定";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnInsertRow;
        private System.Windows.Forms.Button btnRemoveRow;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmKeyword;
        private System.Windows.Forms.DataGridViewComboBoxColumn clmMode;
        private System.Windows.Forms.DataGridViewCheckBoxColumn clmRegEx;
        private System.Windows.Forms.DataGridViewCheckBoxColumn clmCase;
    }
}