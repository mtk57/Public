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
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.chkOutputInfo = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtBeforeSQL
            // 
            this.txtBeforeSQL.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtBeforeSQL.Location = new System.Drawing.Point(12, 36);
            this.txtBeforeSQL.MaxLength = 0;
            this.txtBeforeSQL.Multiline = true;
            this.txtBeforeSQL.Name = "txtBeforeSQL";
            this.txtBeforeSQL.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtBeforeSQL.Size = new System.Drawing.Size(518, 185);
            this.txtBeforeSQL.TabIndex = 4;
            // 
            // txtAfterSQL
            // 
            this.txtAfterSQL.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtAfterSQL.Location = new System.Drawing.Point(12, 36);
            this.txtAfterSQL.MaxLength = 0;
            this.txtAfterSQL.Multiline = true;
            this.txtAfterSQL.Name = "txtAfterSQL";
            this.txtAfterSQL.ReadOnly = true;
            this.txtAfterSQL.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtAfterSQL.Size = new System.Drawing.Size(518, 201);
            this.txtAfterSQL.TabIndex = 3;
            this.txtAfterSQL.TextChanged += new System.EventHandler(this.txtAfterSQL_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "変更前SQL";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 12);
            this.label2.TabIndex = 0;
            this.label2.Text = "変更後SQL";
            // 
            // btnRun
            // 
            this.btnRun.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRun.Location = new System.Drawing.Point(455, 10);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 23);
            this.btnRun.TabIndex = 2;
            this.btnRun.Text = "実行";
            this.btnRun.UseVisualStyleBackColor = true;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.chkOutputInfo);
            this.splitContainer1.Panel1.Controls.Add(this.label1);
            this.splitContainer1.Panel1.Controls.Add(this.btnRun);
            this.splitContainer1.Panel1.Controls.Add(this.txtBeforeSQL);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.label2);
            this.splitContainer1.Panel2.Controls.Add(this.txtAfterSQL);
            this.splitContainer1.Size = new System.Drawing.Size(542, 486);
            this.splitContainer1.SplitterDistance = 233;
            this.splitContainer1.TabIndex = 1;
            // 
            // chkOutputInfo
            // 
            this.chkOutputInfo.AutoSize = true;
            this.chkOutputInfo.Location = new System.Drawing.Point(338, 12);
            this.chkOutputInfo.Name = "chkOutputInfo";
            this.chkOutputInfo.Size = new System.Drawing.Size(96, 16);
            this.chkOutputInfo.TabIndex = 5;
            this.chkOutputInfo.Text = "付加情報出力";
            this.chkOutputInfo.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(542, 486);
            this.Controls.Add(this.splitContainer1);
            this.Name = "MainForm";
            this.Text = "Simple SQL Adjuster";
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.TextBox txtBeforeSQL;
        private System.Windows.Forms.TextBox txtAfterSQL;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.CheckBox chkOutputInfo;
    }
}
