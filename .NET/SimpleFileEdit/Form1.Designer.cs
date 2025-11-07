namespace SimpleFileEdit
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
            this.btnRefDir = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbFolderPath = new System.Windows.Forms.ComboBox();
            this.chkSearchSubDir = new System.Windows.Forms.CheckBox();
            this.btnDeleteComment = new System.Windows.Forms.Button();
            this.cmbTarget = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnDeleteEmptyRow = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnRefDir
            // 
            this.btnRefDir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefDir.Location = new System.Drawing.Point(622, 65);
            this.btnRefDir.Name = "btnRefDir";
            this.btnRefDir.Size = new System.Drawing.Size(45, 23);
            this.btnRefDir.TabIndex = 2;
            this.btnRefDir.Text = "参照";
            this.btnRefDir.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "フォルダパス";
            // 
            // cmbFolderPath
            // 
            this.cmbFolderPath.AllowDrop = true;
            this.cmbFolderPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbFolderPath.FormattingEnabled = true;
            this.cmbFolderPath.Location = new System.Drawing.Point(17, 68);
            this.cmbFolderPath.Name = "cmbFolderPath";
            this.cmbFolderPath.Size = new System.Drawing.Size(599, 20);
            this.cmbFolderPath.TabIndex = 1;
            // 
            // chkSearchSubDir
            // 
            this.chkSearchSubDir.AutoSize = true;
            this.chkSearchSubDir.Checked = true;
            this.chkSearchSubDir.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSearchSubDir.Location = new System.Drawing.Point(22, 93);
            this.chkSearchSubDir.Name = "chkSearchSubDir";
            this.chkSearchSubDir.Size = new System.Drawing.Size(111, 16);
            this.chkSearchSubDir.TabIndex = 3;
            this.chkSearchSubDir.Text = "サブフォルダも対象";
            this.chkSearchSubDir.UseVisualStyleBackColor = true;
            // 
            // btnDeleteComment
            // 
            this.btnDeleteComment.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDeleteComment.Location = new System.Drawing.Point(17, 160);
            this.btnDeleteComment.Name = "btnDeleteComment";
            this.btnDeleteComment.Size = new System.Drawing.Size(87, 23);
            this.btnDeleteComment.TabIndex = 4;
            this.btnDeleteComment.Text = "コメント削除";
            this.btnDeleteComment.UseVisualStyleBackColor = true;
            // 
            // cmbTarget
            // 
            this.cmbTarget.AllowDrop = true;
            this.cmbTarget.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbTarget.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTarget.FormattingEnabled = true;
            this.cmbTarget.Items.AddRange(new object[] {
            "Java"});
            this.cmbTarget.Location = new System.Drawing.Point(52, 12);
            this.cmbTarget.Name = "cmbTarget";
            this.cmbTarget.Size = new System.Drawing.Size(111, 20);
            this.cmbTarget.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 15);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "対象";
            // 
            // btnDeleteEmptyRow
            // 
            this.btnDeleteEmptyRow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDeleteEmptyRow.Location = new System.Drawing.Point(119, 160);
            this.btnDeleteEmptyRow.Name = "btnDeleteEmptyRow";
            this.btnDeleteEmptyRow.Size = new System.Drawing.Size(87, 23);
            this.btnDeleteEmptyRow.TabIndex = 7;
            this.btnDeleteEmptyRow.Text = "空行削除";
            this.btnDeleteEmptyRow.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(679, 226);
            this.Controls.Add(this.btnDeleteEmptyRow);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmbTarget);
            this.Controls.Add(this.btnDeleteComment);
            this.Controls.Add(this.chkSearchSubDir);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbFolderPath);
            this.Controls.Add(this.btnRefDir);
            this.Name = "MainForm";
            this.Text = "Simple File Edit";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnRefDir;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbFolderPath;
        private System.Windows.Forms.CheckBox chkSearchSubDir;
        private System.Windows.Forms.Button btnDeleteComment;
        private System.Windows.Forms.ComboBox cmbTarget;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnDeleteEmptyRow;
    }
}

