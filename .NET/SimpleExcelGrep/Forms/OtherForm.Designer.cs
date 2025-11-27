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
            this.components = new System.ComponentModel.Container();
            this.chkAllShape = new System.Windows.Forms.CheckBox();
            this.chkAllFormula = new System.Windows.Forms.CheckBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProgress = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtKeyword = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.label2 = new System.Windows.Forms.Label();
            this.txtTargetNum = new System.Windows.Forms.TextBox();
            this.chkEnableDeleteByKeyword = new System.Windows.Forms.CheckBox();
            this.rbtnRow = new System.Windows.Forms.RadioButton();
            this.rbtnClm = new System.Windows.Forms.RadioButton();
            this.chkAll = new System.Windows.Forms.CheckBox();
            this.chkFullMatch = new System.Windows.Forms.CheckBox();
            this.chkCaseSensitive = new System.Windows.Forms.CheckBox();
            this.chkWidthSensitive = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // chkAllShape
            // 
            this.chkAllShape.AutoSize = true;
            this.chkAllShape.Location = new System.Drawing.Point(24, 21);
            this.chkAllShape.Name = "chkAllShape";
            this.chkAllShape.Size = new System.Drawing.Size(100, 16);
            this.chkAllShape.TabIndex = 0;
            this.chkAllShape.Text = "全ての図を削除";
            this.chkAllShape.UseVisualStyleBackColor = true;
            // 
            // chkAllFormula
            // 
            this.chkAllFormula.AutoSize = true;
            this.chkAllFormula.Location = new System.Drawing.Point(143, 21);
            this.chkAllFormula.Name = "chkAllFormula";
            this.chkAllFormula.Size = new System.Drawing.Size(133, 16);
            this.chkAllFormula.TabIndex = 1;
            this.chkAllFormula.Text = "全ての数式を値に変換";
            this.chkAllFormula.UseVisualStyleBackColor = true;
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(329, 253);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(49, 26);
            this.btnRun.TabIndex = 4;
            this.btnRun.Text = "実行";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.BtnRun_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(384, 253);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(49, 26);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "中止";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(24, 267);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(299, 12);
            this.progressBar1.TabIndex = 3;
            // 
            // lblProgress
            // 
            this.lblProgress.Location = new System.Drawing.Point(22, 246);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(260, 18);
            this.lblProgress.TabIndex = 2;
            this.lblProgress.Text = "準備完了";
            this.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkWidthSensitive);
            this.groupBox1.Controls.Add(this.chkCaseSensitive);
            this.groupBox1.Controls.Add(this.chkFullMatch);
            this.groupBox1.Controls.Add(this.chkAll);
            this.groupBox1.Controls.Add(this.rbtnClm);
            this.groupBox1.Controls.Add(this.rbtnRow);
            this.groupBox1.Controls.Add(this.chkEnableDeleteByKeyword);
            this.groupBox1.Controls.Add(this.txtTargetNum);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtKeyword);
            this.groupBox1.Location = new System.Drawing.Point(25, 57);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(396, 145);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "キーワードがある列・行を削除";
            // 
            // txtKeyword
            // 
            this.txtKeyword.Location = new System.Drawing.Point(75, 91);
            this.txtKeyword.Name = "txtKeyword";
            this.txtKeyword.Size = new System.Drawing.Size(315, 19);
            this.txtKeyword.TabIndex = 0;
            this.toolTip1.SetToolTip(this.txtKeyword, "カンマ区切りで複数指定可");
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 91);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "キーワード";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 69);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "行番号（列番号）";
            // 
            // txtTargetNum
            // 
            this.txtTargetNum.Location = new System.Drawing.Point(111, 66);
            this.txtTargetNum.Name = "txtTargetNum";
            this.txtTargetNum.Size = new System.Drawing.Size(228, 19);
            this.txtTargetNum.TabIndex = 3;
            this.toolTip1.SetToolTip(this.txtTargetNum, "カンマ区切りで複数指定可。\r\n例1（行番号）：1,2\r\n例2（列番号）：A,B\r\nただし行と列の混在は不可。例：A,3,B\r\n\r\n範囲指定も可。\r\n例1（行番号）" +
        "：1-10\r\n例2（列番号）：A-Z");
            // 
            // chkEnableDeleteByKeyword
            // 
            this.chkEnableDeleteByKeyword.AutoSize = true;
            this.chkEnableDeleteByKeyword.Location = new System.Drawing.Point(18, 27);
            this.chkEnableDeleteByKeyword.Name = "chkEnableDeleteByKeyword";
            this.chkEnableDeleteByKeyword.Size = new System.Drawing.Size(78, 16);
            this.chkEnableDeleteByKeyword.TabIndex = 7;
            this.chkEnableDeleteByKeyword.Text = "有効/無効";
            this.chkEnableDeleteByKeyword.UseVisualStyleBackColor = true;
            this.chkEnableDeleteByKeyword.CheckedChanged += new System.EventHandler(this.ChkEnableDeleteByKeyword_CheckedChanged);
            // 
            // rbtnRow
            // 
            this.rbtnRow.AutoSize = true;
            this.rbtnRow.Checked = true;
            this.rbtnRow.Location = new System.Drawing.Point(118, 27);
            this.rbtnRow.Name = "rbtnRow";
            this.rbtnRow.Size = new System.Drawing.Size(35, 16);
            this.rbtnRow.TabIndex = 8;
            this.rbtnRow.TabStop = true;
            this.rbtnRow.Text = "行";
            this.rbtnRow.UseVisualStyleBackColor = true;
            // 
            // rbtnClm
            // 
            this.rbtnClm.AutoSize = true;
            this.rbtnClm.Location = new System.Drawing.Point(169, 27);
            this.rbtnClm.Name = "rbtnClm";
            this.rbtnClm.Size = new System.Drawing.Size(35, 16);
            this.rbtnClm.TabIndex = 9;
            this.rbtnClm.Text = "列";
            this.rbtnClm.UseVisualStyleBackColor = true;
            // 
            // chkAll
            // 
            this.chkAll.AutoSize = true;
            this.chkAll.Location = new System.Drawing.Point(345, 68);
            this.chkAll.Name = "chkAll";
            this.chkAll.Size = new System.Drawing.Size(45, 16);
            this.chkAll.TabIndex = 10;
            this.chkAll.Text = "全て";
            this.chkAll.UseVisualStyleBackColor = true;
            this.chkAll.CheckedChanged += new System.EventHandler(this.ChkAll_CheckedChanged);
            // 
            // chkFullMatch
            // 
            this.chkFullMatch.AutoSize = true;
            this.chkFullMatch.Location = new System.Drawing.Point(78, 116);
            this.chkFullMatch.Name = "chkFullMatch";
            this.chkFullMatch.Size = new System.Drawing.Size(72, 16);
            this.chkFullMatch.TabIndex = 11;
            this.chkFullMatch.Text = "完全一致";
            this.chkFullMatch.UseVisualStyleBackColor = true;
            // 
            // chkCaseSensitive
            // 
            this.chkCaseSensitive.AutoSize = true;
            this.chkCaseSensitive.Location = new System.Drawing.Point(156, 116);
            this.chkCaseSensitive.Name = "chkCaseSensitive";
            this.chkCaseSensitive.Size = new System.Drawing.Size(72, 16);
            this.chkCaseSensitive.TabIndex = 12;
            this.chkCaseSensitive.Text = "大小区別";
            this.chkCaseSensitive.UseVisualStyleBackColor = true;
            // 
            // chkWidthSensitive
            // 
            this.chkWidthSensitive.AutoSize = true;
            this.chkWidthSensitive.Location = new System.Drawing.Point(234, 116);
            this.chkWidthSensitive.Name = "chkWidthSensitive";
            this.chkWidthSensitive.Size = new System.Drawing.Size(84, 16);
            this.chkWidthSensitive.TabIndex = 13;
            this.chkWidthSensitive.Text = "全半角区別";
            this.chkWidthSensitive.UseVisualStyleBackColor = true;
            // 
            // OtherForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(445, 292);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.chkAllFormula);
            this.Controls.Add(this.chkAllShape);
            this.MaximumSize = new System.Drawing.Size(461, 331);
            this.MinimumSize = new System.Drawing.Size(461, 331);
            this.Name = "OtherForm";
            this.Text = "その他";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
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
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtKeyword;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.TextBox txtTargetNum;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox chkEnableDeleteByKeyword;
        private System.Windows.Forms.RadioButton rbtnRow;
        private System.Windows.Forms.RadioButton rbtnClm;
        private System.Windows.Forms.CheckBox chkAll;
        private System.Windows.Forms.CheckBox chkFullMatch;
        private System.Windows.Forms.CheckBox chkWidthSensitive;
        private System.Windows.Forms.CheckBox chkCaseSensitive;
    }
}
