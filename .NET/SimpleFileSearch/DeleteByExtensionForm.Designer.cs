namespace SimpleFileSearch
{
    partial class DeleteByExtensionForm
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
            this.label5 = new System.Windows.Forms.Label();
            this.txtExtFilter = new System.Windows.Forms.TextBox();
            this.dataGridViewResults = new System.Windows.Forms.DataGridView();
            this.clmExt = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clmCount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnDeleteSelected = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResults)).BeginInit();
            this.SuspendLayout();
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(19, 18);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(41, 12);
            this.label5.TabIndex = 27;
            this.label5.Text = "拡張子";
            // 
            // txtExtFilter
            // 
            this.txtExtFilter.Location = new System.Drawing.Point(21, 33);
            this.txtExtFilter.Name = "txtExtFilter";
            this.txtExtFilter.Size = new System.Drawing.Size(61, 19);
            this.txtExtFilter.TabIndex = 24;
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
            this.clmExt,
            this.clmCount});
            this.dataGridViewResults.Location = new System.Drawing.Point(21, 58);
            this.dataGridViewResults.Name = "dataGridViewResults";
            this.dataGridViewResults.ReadOnly = true;
            this.dataGridViewResults.MultiSelect = true;
            this.dataGridViewResults.RowHeadersVisible = false;
            this.dataGridViewResults.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewResults.Size = new System.Drawing.Size(298, 134);
            this.dataGridViewResults.TabIndex = 21;
            // 
            // clmExt
            // 
            this.clmExt.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmExt.HeaderText = "拡張子";
            this.clmExt.Name = "clmExt";
            this.clmExt.ReadOnly = true;
            // 
            // clmCount
            // 
            this.clmCount.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.clmCount.HeaderText = "ファイル数";
            this.clmCount.Name = "clmCount";
            this.clmCount.ReadOnly = true;
            // 
            // btnDeleteSelected
            // 
            this.btnDeleteSelected.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDeleteSelected.Location = new System.Drawing.Point(89, 201);
            this.btnDeleteSelected.MaximumSize = new System.Drawing.Size(189, 23);
            this.btnDeleteSelected.MinimumSize = new System.Drawing.Size(189, 23);
            this.btnDeleteSelected.Name = "btnDeleteSelected";
            this.btnDeleteSelected.Size = new System.Drawing.Size(189, 23);
            this.btnDeleteSelected.TabIndex = 28;
            this.btnDeleteSelected.Text = "選択行の拡張子のファイルを削除";
            this.btnDeleteSelected.UseVisualStyleBackColor = true;
            // 
            // DeleteByExtensionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(331, 233);
            this.Controls.Add(this.btnDeleteSelected);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtExtFilter);
            this.Controls.Add(this.dataGridViewResults);
            this.Name = "DeleteByExtensionForm";
            this.Text = "拡張子で削除";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResults)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtExtFilter;
        private System.Windows.Forms.DataGridView dataGridViewResults;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmExt;
        private System.Windows.Forms.DataGridViewTextBoxColumn clmCount;
        private System.Windows.Forms.Button btnDeleteSelected;
    }
}
