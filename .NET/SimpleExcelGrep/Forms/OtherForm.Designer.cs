namespace SimpleExcelGrep.Forms
{
    partial class OtherForm
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
            this.chkAllShape = new System.Windows.Forms.CheckBox();
            this.chkAllFormula = new System.Windows.Forms.CheckBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // chkAllShape
            // 
            this.chkAllShape.AutoSize = true;
            this.chkAllShape.Location = new System.Drawing.Point(24, 13);
            this.chkAllShape.Name = "chkAllShape";
            this.chkAllShape.Size = new System.Drawing.Size(67, 16);
            this.chkAllShape.TabIndex = 0;
            this.chkAllShape.Text = "全ての図";
            this.chkAllShape.UseVisualStyleBackColor = true;
            // 
            // chkAllFormula
            // 
            this.chkAllFormula.AutoSize = true;
            this.chkAllFormula.Location = new System.Drawing.Point(113, 13);
            this.chkAllFormula.Name = "chkAllFormula";
            this.chkAllFormula.Size = new System.Drawing.Size(79, 16);
            this.chkAllFormula.TabIndex = 1;
            this.chkAllFormula.Text = "全ての数式";
            this.chkAllFormula.UseVisualStyleBackColor = true;
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(24, 152);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(102, 23);
            this.btnRun.TabIndex = 2;
            this.btnRun.Text = "実行";
            this.btnRun.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(149, 152);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(85, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "中止";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(24, 136);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(210, 10);
            this.progressBar1.TabIndex = 4;
            // 
            // OtherForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(300, 204);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.chkAllFormula);
            this.Controls.Add(this.chkAllShape);
            this.MaximumSize = new System.Drawing.Size(316, 243);
            this.MinimumSize = new System.Drawing.Size(316, 243);
            this.Name = "OtherForm";
            this.Text = "その他";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkAllShape;
        private System.Windows.Forms.CheckBox chkAllFormula;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}