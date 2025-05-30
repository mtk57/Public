using System.Drawing;
using System.Windows.Forms;

namespace ER2SpreadTool
{
    partial class MainForm
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
            label1 = new Label();
            txtFilePath = new TextBox();
            btnBrowse = new Button();
            label2 = new Label();
            txtSheetName = new TextBox();
            btnProcess = new Button();
            label3 = new Label();
            txtResults = new TextBox();
            lblStatus = new Label();
            SuspendLayout();

                        // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(23, 25);
            label1.Name = "label1";
            label1.Size = new Size(87, 15);
            label1.TabIndex = 0;
            label1.Text = "Excelファイルパス";
            // 
            // txtFilePath
            // 
            txtFilePath.Location = new Point(23, 43);
            txtFilePath.Name = "txtFilePath";
            txtFilePath.Size = new Size(526, 23);
            txtFilePath.TabIndex = 1;
            // 
            // btnBrowse
            // 
            btnBrowse.Location = new Point(555, 43);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(75, 23);
            btnBrowse.TabIndex = 2;
            btnBrowse.Text = "参照";
            btnBrowse.UseVisualStyleBackColor = true;
            btnBrowse.Click += btnBrowse_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(23, 103);
            label2.Name = "label2";
            label2.Size = new Size(45, 15);
            label2.TabIndex = 3;
            label2.Text = "シート名";
            // 
            // txtSheetName
            // 
            txtSheetName.Location = new Point(23, 121);
            txtSheetName.Name = "txtSheetName";
            txtSheetName.Size = new Size(526, 23);
            txtSheetName.TabIndex = 4;
            // 
            // btnProcess
            // 
            btnProcess.Location = new Point(277, 296);
            btnProcess.Name = "btnProcess";
            btnProcess.Size = new Size(75, 23);
            btnProcess.TabIndex = 5;
            btnProcess.Text = "処理開始";
            btnProcess.UseVisualStyleBackColor = true;
            btnProcess.Click += btnProcess_Click;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(23, 170);
            label3.Name = "label3";
            label3.Size = new Size(55, 15);
            label3.TabIndex = 6;
            label3.Text = "処理結果";
            // 
            // txtResults
            // 
            txtResults.Location = new Point(23, 187);
            txtResults.Multiline = true;
            txtResults.Name = "txtResults";
            txtResults.ReadOnly = true;
            txtResults.ScrollBars = ScrollBars.Both;
            txtResults.Size = new Size(607, 87);
            txtResults.TabIndex = 7;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(23, 347);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(55, 15);
            lblStatus.TabIndex = 8;
            lblStatus.Text = "準備完了";

            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(656, 371);
            Controls.Add(lblStatus);
            Controls.Add(txtResults);
            Controls.Add(label3);
            Controls.Add(btnProcess);
            Controls.Add(txtSheetName);
            Controls.Add(label2);
            Controls.Add(btnBrowse);
            Controls.Add(txtFilePath);
            Controls.Add(label1);

            Name = "MainForm";
            Text = "ER図スプレッドツール";
            ResumeLayout(false);
            PerformLayout();
        }

        private Label label1;
        private TextBox txtFilePath;
        private Button btnBrowse;
        private Label label2;
        private TextBox txtSheetName;
        private Button btnProcess;
        private Label label3;
        private TextBox txtResults;
        private Label lblStatus;

        #endregion
    }
}