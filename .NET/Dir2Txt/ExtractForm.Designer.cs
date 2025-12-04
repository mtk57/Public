namespace Dir2Txt
{
    partial class ExtractForm
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
            this.btnRun = new System.Windows.Forms.Button();
            this.btnRefExtractDirPath = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtExtractDirPath = new System.Windows.Forms.TextBox();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(225, 62);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(44, 23);
            this.btnRun.TabIndex = 9;
            this.btnRun.Text = "実行";
            this.btnRun.UseVisualStyleBackColor = true;
            // 
            // btnRefExtractDirPath
            // 
            this.btnRefExtractDirPath.Location = new System.Drawing.Point(437, 24);
            this.btnRefExtractDirPath.Name = "btnRefExtractDirPath";
            this.btnRefExtractDirPath.Size = new System.Drawing.Size(44, 23);
            this.btnRefExtractDirPath.TabIndex = 12;
            this.btnRefExtractDirPath.Text = "ref";
            this.btnRefExtractDirPath.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 12);
            this.label2.TabIndex = 11;
            this.label2.Text = "復元フォルダパス";
            // 
            // txtExtractDirPath
            // 
            this.txtExtractDirPath.Location = new System.Drawing.Point(30, 27);
            this.txtExtractDirPath.Name = "txtExtractDirPath";
            this.txtExtractDirPath.Size = new System.Drawing.Size(400, 19);
            this.txtExtractDirPath.TabIndex = 10;
            // 
            // txtOutput
            // 
            this.txtOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtOutput.Location = new System.Drawing.Point(19, 98);
            this.txtOutput.MaxLength = 0;
            this.txtOutput.Multiline = true;
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtOutput.Size = new System.Drawing.Size(476, 307);
            this.txtOutput.TabIndex = 13;
            // 
            // ExtractForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(514, 427);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.btnRefExtractDirPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtExtractDirPath);
            this.Controls.Add(this.btnRun);
            this.Name = "ExtractForm";
            this.Text = "復元";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Button btnRefExtractDirPath;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtExtractDirPath;
        private System.Windows.Forms.TextBox txtOutput;
    }
}