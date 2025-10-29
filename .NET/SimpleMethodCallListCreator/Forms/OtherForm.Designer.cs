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
            this.SuspendLayout();
            // 
            // btnMethodList
            // 
            this.btnMethodList.Location = new System.Drawing.Point(140, 28);
            this.btnMethodList.Name = "btnMethodList";
            this.btnMethodList.Size = new System.Drawing.Size(128, 29);
            this.btnMethodList.TabIndex = 0;
            this.btnMethodList.Text = "メソッドリスト作成";
            this.btnMethodList.UseVisualStyleBackColor = true;
            // 
            // btnInsertTagJump
            // 
            this.btnInsertTagJump.Location = new System.Drawing.Point(140, 79);
            this.btnInsertTagJump.Name = "btnInsertTagJump";
            this.btnInsertTagJump.Size = new System.Drawing.Size(128, 29);
            this.btnInsertTagJump.TabIndex = 1;
            this.btnInsertTagJump.Text = "タグジャンプ埋め込み";
            this.btnInsertTagJump.UseVisualStyleBackColor = true;
            this.btnInsertTagJump.Click += new System.EventHandler(this.btnInsertTagJump_Click);
            // 
            // OtherForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(420, 209);
            this.Controls.Add(this.btnInsertTagJump);
            this.Controls.Add(this.btnMethodList);
            this.Name = "OtherForm";
            this.Text = "Other";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnMethodList;
        private System.Windows.Forms.Button btnInsertTagJump;
    }
}