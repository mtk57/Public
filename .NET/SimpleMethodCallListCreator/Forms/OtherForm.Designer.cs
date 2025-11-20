namespace SimpleMethodCallListCreator
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
            this.btnMethodList = new System.Windows.Forms.Button();
            this.btnInsertTagJump = new System.Windows.Forms.Button();
            this.btnInsertTagJumpWithRowNum = new System.Windows.Forms.Button();
            this.btnMethodListWithRowNum = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnCollectFiles = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnMethodList
            // 
            this.btnMethodList.Location = new System.Drawing.Point(42, 31);
            this.btnMethodList.Name = "btnMethodList";
            this.btnMethodList.Size = new System.Drawing.Size(128, 29);
            this.btnMethodList.TabIndex = 0;
            this.btnMethodList.Text = "メソッドリスト作成";
            this.btnMethodList.UseVisualStyleBackColor = true;
            // 
            // btnInsertTagJump
            // 
            this.btnInsertTagJump.Location = new System.Drawing.Point(42, 71);
            this.btnInsertTagJump.Name = "btnInsertTagJump";
            this.btnInsertTagJump.Size = new System.Drawing.Size(128, 29);
            this.btnInsertTagJump.TabIndex = 1;
            this.btnInsertTagJump.Text = "タグジャンプ埋め込み";
            this.btnInsertTagJump.UseVisualStyleBackColor = true;
            this.btnInsertTagJump.Click += new System.EventHandler(this.btnInsertTagJump_Click);
            // 
            // btnInsertTagJumpWithRowNum
            // 
            this.btnInsertTagJumpWithRowNum.Location = new System.Drawing.Point(50, 71);
            this.btnInsertTagJumpWithRowNum.Name = "btnInsertTagJumpWithRowNum";
            this.btnInsertTagJumpWithRowNum.Size = new System.Drawing.Size(128, 29);
            this.btnInsertTagJumpWithRowNum.TabIndex = 3;
            this.btnInsertTagJumpWithRowNum.Text = "タグジャンプ埋め込み";
            this.btnInsertTagJumpWithRowNum.UseVisualStyleBackColor = true;
            this.btnInsertTagJumpWithRowNum.Click += new System.EventHandler(this.btnInsertTagJumpWithRowNum_Click);
            // 
            // btnMethodListWithRowNum
            // 
            this.btnMethodListWithRowNum.Location = new System.Drawing.Point(50, 31);
            this.btnMethodListWithRowNum.Name = "btnMethodListWithRowNum";
            this.btnMethodListWithRowNum.Size = new System.Drawing.Size(128, 29);
            this.btnMethodListWithRowNum.TabIndex = 2;
            this.btnMethodListWithRowNum.Text = "メソッドリスト作成";
            this.btnMethodListWithRowNum.UseVisualStyleBackColor = true;
            this.btnMethodListWithRowNum.Click += new System.EventHandler(this.btnMethodListWithRowNum_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnCollectFiles);
            this.groupBox1.Controls.Add(this.btnMethodList);
            this.groupBox1.Controls.Add(this.btnInsertTagJump);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(222, 212);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "ファイルパス+シグネチャ+メソッドリストパス";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnMethodListWithRowNum);
            this.groupBox2.Controls.Add(this.btnInsertTagJumpWithRowNum);
            this.groupBox2.Location = new System.Drawing.Point(255, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(219, 135);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "ファイルパス＋行番号";
            // 
            // btnCollectFiles
            // 
            this.btnCollectFiles.Location = new System.Drawing.Point(42, 132);
            this.btnCollectFiles.Name = "btnCollectFiles";
            this.btnCollectFiles.Size = new System.Drawing.Size(128, 29);
            this.btnCollectFiles.TabIndex = 2;
            this.btnCollectFiles.Text = "ファイル収集";
            this.btnCollectFiles.UseVisualStyleBackColor = true;
            // 
            // OtherForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(510, 236);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "OtherForm";
            this.Text = "その他の機能";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnMethodList;
        private System.Windows.Forms.Button btnInsertTagJump;
        private System.Windows.Forms.Button btnInsertTagJumpWithRowNum;
        private System.Windows.Forms.Button btnMethodListWithRowNum;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnCollectFiles;
    }
}