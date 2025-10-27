namespace SimpleSqlAdjuster
{
    partial class MainForm
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose ( bool disposing )
        {
            if ( disposing && ( components != null ) )
            {
                components.Dispose();
            }
            base.Dispose( disposing );
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent ()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.txtBeforeSQL = new System.Windows.Forms.TextBox();
            this.txtAfterSQL = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnRun = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtBeforeSQL
            // 
            this.txtBeforeSQL.Location = new System.Drawing.Point(31, 46);
            this.txtBeforeSQL.MaxLength = 0;
            this.txtBeforeSQL.Multiline = true;
            this.txtBeforeSQL.Name = "txtBeforeSQL";
            this.txtBeforeSQL.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtBeforeSQL.Size = new System.Drawing.Size(491, 187);
            this.txtBeforeSQL.TabIndex = 0;
            this.txtBeforeSQL.WordWrap = false;
            // 
            // txtAfterSQL
            // 
            this.txtAfterSQL.Location = new System.Drawing.Point(27, 283);
            this.txtAfterSQL.MaxLength = 0;
            this.txtAfterSQL.Multiline = true;
            this.txtAfterSQL.Name = "txtAfterSQL";
            this.txtAfterSQL.ReadOnly = true;
            this.txtAfterSQL.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtAfterSQL.Size = new System.Drawing.Size(496, 182);
            this.txtAfterSQL.TabIndex = 1;
            this.txtAfterSQL.WordWrap = false;
            this.txtAfterSQL.TextChanged += new System.EventHandler(this.txtAfterSQL_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(31, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "変更前SQL";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 249);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "変更後SQL";
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(220, 12);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 23);
            this.btnRun.TabIndex = 4;
            this.btnRun.Text = "実行";
            this.btnRun.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(542, 486);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtAfterSQL);
            this.Controls.Add(this.txtBeforeSQL);
            this.Name = "MainForm";
            this.Text = "Simple SQL Adjuster";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.TextBox txtBeforeSQL;
        private System.Windows.Forms.TextBox txtAfterSQL;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnRun;
    }
}

