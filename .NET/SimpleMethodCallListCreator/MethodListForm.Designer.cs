namespace SimpleMethodCallListCreator
{
    partial class MethodListForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.cmbDirPath = new System.Windows.Forms.ComboBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.cmbExt = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "対象フォルダパス";
            // 
            // cmbDirPath
            // 
            this.cmbDirPath.AllowDrop = true;
            this.cmbDirPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbDirPath.FormattingEnabled = true;
            this.cmbDirPath.Location = new System.Drawing.Point(25, 39);
            this.cmbDirPath.Name = "cmbDirPath";
            this.cmbDirPath.Size = new System.Drawing.Size(638, 20);
            this.cmbDirPath.TabIndex = 4;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowse.Location = new System.Drawing.Point(669, 39);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(45, 23);
            this.btnBrowse.TabIndex = 5;
            this.btnBrowse.Text = "参照";
            this.btnBrowse.UseVisualStyleBackColor = true;
            // 
            // cmbExt
            // 
            this.cmbExt.AllowDrop = true;
            this.cmbExt.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbExt.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbExt.FormattingEnabled = true;
            this.cmbExt.Items.AddRange(new object[] {
            "Java"});
            this.cmbExt.Location = new System.Drawing.Point(25, 101);
            this.cmbExt.Name = "cmbExt";
            this.cmbExt.Size = new System.Drawing.Size(81, 20);
            this.cmbExt.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 86);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 7;
            this.label2.Text = "対象拡張子";
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(300, 128);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(103, 29);
            this.button1.TabIndex = 8;
            this.button1.Text = "メソッドリスト作成";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // MethodListForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(736, 183);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmbExt);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbDirPath);
            this.Controls.Add(this.btnBrowse);
            this.MaximumSize = new System.Drawing.Size(752, 222);
            this.MinimumSize = new System.Drawing.Size(752, 222);
            this.Name = "MethodListForm";
            this.Text = "Method List";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbDirPath;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.ComboBox cmbExt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
    }
}