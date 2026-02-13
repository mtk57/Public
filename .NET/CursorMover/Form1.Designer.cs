namespace CursorMover
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
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.textCursorMoveTimeSec = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnStop = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.label2 = new System.Windows.Forms.Label();
            this.lblElapsedTime = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtElapsedTimeBefore = new System.Windows.Forms.TextBox();
            this.btnSaveElapsedTimeBefore = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.txtMemo = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // textCursorMoveTimeSec
            // 
            this.textCursorMoveTimeSec.Location = new System.Drawing.Point(173, 16);
            this.textCursorMoveTimeSec.Name = "textCursorMoveTimeSec";
            this.textCursorMoveTimeSec.Size = new System.Drawing.Size(55, 19);
            this.textCursorMoveTimeSec.TabIndex = 1;
            this.textCursorMoveTimeSec.Text = "5";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(139, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "カーソル自動移動間隔（秒）";
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(45, 163);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(79, 25);
            this.btnStart.TabIndex = 2;
            this.btnStart.Text = "START";
            this.btnStart.UseVisualStyleBackColor = true;
            // 
            // btnStop
            // 
            this.btnStop.Enabled = false;
            this.btnStop.Location = new System.Drawing.Point(148, 163);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new System.Drawing.Size(79, 25);
            this.btnStop.TabIndex = 3;
            this.btnStop.Text = "STOP";
            this.btnStop.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "経過時間";
            // 
            // lblElapsedTime
            // 
            this.lblElapsedTime.AutoSize = true;
            this.lblElapsedTime.Location = new System.Drawing.Point(119, 60);
            this.lblElapsedTime.Name = "lblElapsedTime";
            this.lblElapsedTime.Size = new System.Drawing.Size(55, 12);
            this.lblElapsedTime.TabIndex = 5;
            this.lblElapsedTime.Text = "HH:mm:ss";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 91);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(89, 12);
            this.label4.TabIndex = 6;
            this.label4.Text = "経過時間（前回）";
            // 
            // txtElapsedTimeBefore
            // 
            this.txtElapsedTimeBefore.Location = new System.Drawing.Point(121, 88);
            this.txtElapsedTimeBefore.Name = "txtElapsedTimeBefore";
            this.txtElapsedTimeBefore.Size = new System.Drawing.Size(78, 19);
            this.txtElapsedTimeBefore.TabIndex = 7;
            // 
            // btnSaveElapsedTimeBefore
            // 
            this.btnSaveElapsedTimeBefore.Enabled = false;
            this.btnSaveElapsedTimeBefore.Location = new System.Drawing.Point(216, 88);
            this.btnSaveElapsedTimeBefore.Name = "btnSaveElapsedTimeBefore";
            this.btnSaveElapsedTimeBefore.Size = new System.Drawing.Size(44, 19);
            this.btnSaveElapsedTimeBefore.TabIndex = 8;
            this.btnSaveElapsedTimeBefore.Text = "保持";
            this.btnSaveElapsedTimeBefore.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 126);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(22, 12);
            this.label5.TabIndex = 9;
            this.label5.Text = "メモ";
            // 
            // txtMemo
            // 
            this.txtMemo.Location = new System.Drawing.Point(121, 119);
            this.txtMemo.Name = "txtMemo";
            this.txtMemo.Size = new System.Drawing.Size(139, 19);
            this.txtMemo.TabIndex = 10;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 200);
            this.Controls.Add(this.txtMemo);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnSaveElapsedTimeBefore);
            this.Controls.Add(this.txtElapsedTimeBefore);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblElapsedTime);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnStop);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textCursorMoveTimeSec);
            this.Name = "MainForm";
            this.Text = "Cursor Mover";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textCursorMoveTimeSec;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnStop;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblElapsedTime;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtElapsedTimeBefore;
        private System.Windows.Forms.Button btnSaveElapsedTimeBefore;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtMemo;
    }
}