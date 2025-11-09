namespace SimpleMethodCallListCreator.Forms
{
    partial class InsertTagJumpForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.btnRefStartSrcFilePath = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btnRefMethodListPath = new System.Windows.Forms.Button();
            this.txtMethodListPath = new System.Windows.Forms.TextBox();
            this.txtStartSrcFilePath = new System.Windows.Forms.TextBox();
            this.txtStartMethod = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtSrcRootDirPath = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnRefSrcRootDirPath = new System.Windows.Forms.Button();
            this.txtTagJumpPrefix = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.lblFailed = new System.Windows.Forms.Label();
            this.chkAllMethodMode = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(366, 321);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 33);
            this.btnRun.TabIndex = 10;
            this.btnRun.Text = "実行";
            this.btnRun.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(45, 123);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 12);
            this.label1.TabIndex = 6;
            this.label1.Text = "開始ソースファイルパス";
            // 
            // btnRefStartSrcFilePath
            // 
            this.btnRefStartSrcFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefStartSrcFilePath.Location = new System.Drawing.Point(722, 138);
            this.btnRefStartSrcFilePath.Name = "btnRefStartSrcFilePath";
            this.btnRefStartSrcFilePath.Size = new System.Drawing.Size(45, 23);
            this.btnRefStartSrcFilePath.TabIndex = 8;
            this.btnRefStartSrcFilePath.Text = "参照";
            this.btnRefStartSrcFilePath.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(45, 27);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(81, 12);
            this.label2.TabIndex = 11;
            this.label2.Text = "メソッドリストパス";
            // 
            // btnRefMethodListPath
            // 
            this.btnRefMethodListPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefMethodListPath.Location = new System.Drawing.Point(722, 42);
            this.btnRefMethodListPath.Name = "btnRefMethodListPath";
            this.btnRefMethodListPath.Size = new System.Drawing.Size(45, 23);
            this.btnRefMethodListPath.TabIndex = 13;
            this.btnRefMethodListPath.Text = "参照";
            this.btnRefMethodListPath.UseVisualStyleBackColor = true;
            // 
            // txtMethodListPath
            // 
            this.txtMethodListPath.Location = new System.Drawing.Point(47, 42);
            this.txtMethodListPath.Name = "txtMethodListPath";
            this.txtMethodListPath.Size = new System.Drawing.Size(669, 19);
            this.txtMethodListPath.TabIndex = 14;
            // 
            // txtStartSrcFilePath
            // 
            this.txtStartSrcFilePath.Location = new System.Drawing.Point(47, 142);
            this.txtStartSrcFilePath.Name = "txtStartSrcFilePath";
            this.txtStartSrcFilePath.Size = new System.Drawing.Size(669, 19);
            this.txtStartSrcFilePath.TabIndex = 15;
            // 
            // txtStartMethod
            // 
            this.txtStartMethod.Location = new System.Drawing.Point(47, 206);
            this.txtStartMethod.Name = "txtStartMethod";
            this.txtStartMethod.Size = new System.Drawing.Size(669, 19);
            this.txtStartMethod.TabIndex = 16;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(45, 187);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(74, 12);
            this.label3.TabIndex = 17;
            this.label3.Text = "開始メソッド名";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(154, 187);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(318, 12);
            this.label4.TabIndex = 18;
            this.label4.Text = "※同名メソッドがある場合はメソッドリストのシグネチャを記載すること";
            // 
            // txtSrcRootDirPath
            // 
            this.txtSrcRootDirPath.Enabled = false;
            this.txtSrcRootDirPath.Location = new System.Drawing.Point(46, 92);
            this.txtSrcRootDirPath.Name = "txtSrcRootDirPath";
            this.txtSrcRootDirPath.Size = new System.Drawing.Size(669, 19);
            this.txtSrcRootDirPath.TabIndex = 21;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(44, 73);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(115, 12);
            this.label5.TabIndex = 19;
            this.label5.Text = "ソースルートフォルダパス";
            // 
            // btnRefSrcRootDirPath
            // 
            this.btnRefSrcRootDirPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefSrcRootDirPath.Enabled = false;
            this.btnRefSrcRootDirPath.Location = new System.Drawing.Point(721, 88);
            this.btnRefSrcRootDirPath.Name = "btnRefSrcRootDirPath";
            this.btnRefSrcRootDirPath.Size = new System.Drawing.Size(45, 23);
            this.btnRefSrcRootDirPath.TabIndex = 20;
            this.btnRefSrcRootDirPath.Text = "参照";
            this.btnRefSrcRootDirPath.UseVisualStyleBackColor = true;
            // 
            // txtTagJumpPrefix
            // 
            this.txtTagJumpPrefix.Location = new System.Drawing.Point(46, 273);
            this.txtTagJumpPrefix.Name = "txtTagJumpPrefix";
            this.txtTagJumpPrefix.Size = new System.Drawing.Size(175, 19);
            this.txtTagJumpPrefix.TabIndex = 23;
            this.txtTagJumpPrefix.Text = "//@ ";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(45, 258);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(177, 12);
            this.label7.TabIndex = 24;
            this.label7.Text = "タグジャンプ情報の前につける文字列";
            // 
            // lblFailed
            // 
            this.lblFailed.AutoSize = true;
            this.lblFailed.ForeColor = System.Drawing.Color.Red;
            this.lblFailed.Location = new System.Drawing.Point(459, 331);
            this.lblFailed.Name = "lblFailed";
            this.lblFailed.Size = new System.Drawing.Size(88, 12);
            this.lblFailed.TabIndex = 25;
            this.lblFailed.Text = "メソッド特定失敗:";
            // 
            // chkAllMethodMode
            // 
            this.chkAllMethodMode.AutoSize = true;
            this.chkAllMethodMode.Location = new System.Drawing.Point(46, 338);
            this.chkAllMethodMode.Name = "chkAllMethodMode";
            this.chkAllMethodMode.Size = new System.Drawing.Size(97, 16);
            this.chkAllMethodMode.TabIndex = 26;
            this.chkAllMethodMode.Text = "全メソッドモード";
            this.chkAllMethodMode.UseVisualStyleBackColor = true;
            // 
            // InsertTagJumpForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 383);
            this.Controls.Add(this.chkAllMethodMode);
            this.Controls.Add(this.lblFailed);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtTagJumpPrefix);
            this.Controls.Add(this.txtSrcRootDirPath);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnRefSrcRootDirPath);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtStartMethod);
            this.Controls.Add(this.txtStartSrcFilePath);
            this.Controls.Add(this.txtMethodListPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnRefMethodListPath);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnRefStartSrcFilePath);
            this.Name = "InsertTagJumpForm";
            this.Text = "Insert Tag Jump";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnRefStartSrcFilePath;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnRefMethodListPath;
        private System.Windows.Forms.TextBox txtMethodListPath;
        private System.Windows.Forms.TextBox txtStartSrcFilePath;
        private System.Windows.Forms.TextBox txtStartMethod;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtSrcRootDirPath;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnRefSrcRootDirPath;
        private System.Windows.Forms.TextBox txtTagJumpPrefix;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label lblFailed;
        private System.Windows.Forms.CheckBox chkAllMethodMode;
    }
}