namespace SimpleFileSearch
{
    partial class DeleteMethodForm
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
            this.txtKeywordForMethodSignature = new System.Windows.Forms.TextBox();
            this.chkEnabledRegExForMethodSignature = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(269, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "メソッドシグネチャ内に次のキーワードがあった場合に削除";
            // 
            // txtKeywordForMethodSignature
            // 
            this.txtKeywordForMethodSignature.Location = new System.Drawing.Point(26, 44);
            this.txtKeywordForMethodSignature.Name = "txtKeywordForMethodSignature";
            this.txtKeywordForMethodSignature.Size = new System.Drawing.Size(276, 19);
            this.txtKeywordForMethodSignature.TabIndex = 1;
            // 
            // chkEnabledRegExForMethodSignature
            // 
            this.chkEnabledRegExForMethodSignature.AutoSize = true;
            this.chkEnabledRegExForMethodSignature.Location = new System.Drawing.Point(26, 70);
            this.chkEnabledRegExForMethodSignature.Name = "chkEnabledRegExForMethodSignature";
            this.chkEnabledRegExForMethodSignature.Size = new System.Drawing.Size(72, 16);
            this.chkEnabledRegExForMethodSignature.TabIndex = 2;
            this.chkEnabledRegExForMethodSignature.Text = "正規表現";
            this.chkEnabledRegExForMethodSignature.UseVisualStyleBackColor = true;
            // 
            // DeleteMethodForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(376, 148);
            this.Controls.Add(this.chkEnabledRegExForMethodSignature);
            this.Controls.Add(this.txtKeywordForMethodSignature);
            this.Controls.Add(this.label1);
            this.Name = "DeleteMethodForm";
            this.Text = "Delete Method Form";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtKeywordForMethodSignature;
        private System.Windows.Forms.CheckBox chkEnabledRegExForMethodSignature;
    }
}