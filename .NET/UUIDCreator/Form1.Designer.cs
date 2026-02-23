namespace UUIDCreator
{
    partial class Form1
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
            this.txtInputFilePath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnRefInputFilePath = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtOutputDirPath = new System.Windows.Forms.TextBox();
            this.btnRefOutputDirPath = new System.Windows.Forms.Button();
            this.btnCreate = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // txtInputFilePath
            // 
            this.txtInputFilePath.Location = new System.Drawing.Point(25, 37);
            this.txtInputFilePath.Name = "txtInputFilePath";
            this.txtInputFilePath.Size = new System.Drawing.Size(444, 19);
            this.txtInputFilePath.TabIndex = 0;
            this.toolTip1.SetToolTip(this.txtInputFilePath, "ドラッグアンドドロップ可");
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "件数ファイルパス";
            // 
            // btnRefInputFilePath
            // 
            this.btnRefInputFilePath.Location = new System.Drawing.Point(476, 37);
            this.btnRefInputFilePath.Name = "btnRefInputFilePath";
            this.btnRefInputFilePath.Size = new System.Drawing.Size(44, 23);
            this.btnRefInputFilePath.TabIndex = 2;
            this.btnRefInputFilePath.Text = "参照";
            this.btnRefInputFilePath.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "出力フォルダパス";
            // 
            // txtOutputDirPath
            // 
            this.txtOutputDirPath.Location = new System.Drawing.Point(25, 113);
            this.txtOutputDirPath.Name = "txtOutputDirPath";
            this.txtOutputDirPath.Size = new System.Drawing.Size(444, 19);
            this.txtOutputDirPath.TabIndex = 3;
            this.toolTip1.SetToolTip(this.txtOutputDirPath, "ドラッグアンドドロップ可");
            // 
            // btnRefOutputDirPath
            // 
            this.btnRefOutputDirPath.Location = new System.Drawing.Point(476, 113);
            this.btnRefOutputDirPath.Name = "btnRefOutputDirPath";
            this.btnRefOutputDirPath.Size = new System.Drawing.Size(44, 23);
            this.btnRefOutputDirPath.TabIndex = 5;
            this.btnRefOutputDirPath.Text = "参照";
            this.btnRefOutputDirPath.UseVisualStyleBackColor = true;
            // 
            // btnCreate
            // 
            this.btnCreate.Location = new System.Drawing.Point(220, 166);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(44, 23);
            this.btnCreate.TabIndex = 6;
            this.btnCreate.Text = "作成";
            this.btnCreate.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(532, 219);
            this.Controls.Add(this.btnCreate);
            this.Controls.Add(this.btnRefOutputDirPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtOutputDirPath);
            this.Controls.Add(this.btnRefInputFilePath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtInputFilePath);
            this.Name = "Form1";
            this.Text = "UUID Creator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtInputFilePath;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnRefInputFilePath;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtOutputDirPath;
        private System.Windows.Forms.Button btnRefOutputDirPath;
        private System.Windows.Forms.Button btnCreate;
    }
}

