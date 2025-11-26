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
            this.lblProgress = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // chkAllShape
            // 
            this.chkAllShape.AutoSize = true;
            this.chkAllShape.Location = new System.Drawing.Point(24, 21);
            this.chkAllShape.Name = "chkAllShape";
            this.chkAllShape.Size = new System.Drawing.Size(67, 16);
            this.chkAllShape.TabIndex = 0;
            this.chkAllShape.Text = "全ての図";
            this.chkAllShape.UseVisualStyleBackColor = true;
            // 
            // chkAllFormula
            // 
            this.chkAllFormula.AutoSize = true;
            this.chkAllFormula.Location = new System.Drawing.Point(117, 21);
            this.chkAllFormula.Name = "chkAllFormula";
            this.chkAllFormula.Size = new System.Drawing.Size(79, 16);
            this.chkAllFormula.TabIndex = 1;
            this.chkAllFormula.Text = "全ての数式";
            this.chkAllFormula.UseVisualStyleBackColor = true;
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(24, 124);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(120, 26);
            this.btnRun.TabIndex = 4;
            this.btnRun.Text = "実行";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.BtnRun_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(162, 124);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(120, 26);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "中止";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(24, 92);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(258, 12);
            this.progressBar1.TabIndex = 3;
            // 
            // lblProgress
            // 
            this.lblProgress.Location = new System.Drawing.Point(22, 71);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(260, 18);
            this.lblProgress.TabIndex = 2;
            this.lblProgress.Text = "準備完了";
            this.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // OtherForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(306, 168);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.chkAllFormula);
            this.Controls.Add(this.chkAllShape);
            this.MaximumSize = new System.Drawing.Size(322, 207);
            this.MinimumSize = new System.Drawing.Size(322, 207);
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
        private System.Windows.Forms.Label lblProgress;
    }
}
