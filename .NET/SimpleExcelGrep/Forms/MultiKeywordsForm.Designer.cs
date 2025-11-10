namespace SimpleExcelGrep.Forms
{
    partial class MultiKeywordsForm
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
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnApply = new System.Windows.Forms.Button();
            this.txtMultiKeywords = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnCancel.Location = new System.Drawing.Point(196, 193);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(69, 29);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "キャンセル";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnApply
            // 
            this.btnApply.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnApply.Location = new System.Drawing.Point(111, 193);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(69, 29);
            this.btnApply.TabIndex = 4;
            this.btnApply.Text = "検索";
            this.btnApply.UseVisualStyleBackColor = true;
            // 
            // txtMultiKeywords
            // 
            this.txtMultiKeywords.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtMultiKeywords.Location = new System.Drawing.Point(12, 12);
            this.txtMultiKeywords.MaxLength = 0;
            this.txtMultiKeywords.Multiline = true;
            this.txtMultiKeywords.Name = "txtMultiKeywords";
            this.txtMultiKeywords.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtMultiKeywords.Size = new System.Drawing.Size(368, 162);
            this.txtMultiKeywords.TabIndex = 3;
            // 
            // MultiKeywordsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(392, 244);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.txtMultiKeywords);
            this.Name = "MultiKeywordsForm";
            this.Text = "複数キーワード指定";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.TextBox txtMultiKeywords;
    }
}