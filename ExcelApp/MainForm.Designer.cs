namespace ExcelReportApplication
{
    partial class MainForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_run = new System.Windows.Forms.Button();
            this.MsgWindow = new System.Windows.Forms.TextBox();
            this.btnSelectBugFile = new System.Windows.Forms.Button();
            this.txtBugFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtTCFile = new System.Windows.Forms.TextBox();
            this.btnSelectTCFile = new System.Windows.Forms.Button();
            this.btnSelectReportFile = new System.Windows.Forms.Button();
            this.txtReportFile = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btn_run
            // 
            this.btn_run.Location = new System.Drawing.Point(435, 123);
            this.btn_run.Name = "btn_run";
            this.btn_run.Size = new System.Drawing.Size(67, 23);
            this.btn_run.TabIndex = 0;
            this.btn_run.Text = "Run";
            this.btn_run.UseVisualStyleBackColor = true;
            this.btn_run.Click += new System.EventHandler(this.button1_Click);
            // 
            // MsgWindow
            // 
            this.MsgWindow.Location = new System.Drawing.Point(12, 152);
            this.MsgWindow.Multiline = true;
            this.MsgWindow.Name = "MsgWindow";
            this.MsgWindow.ReadOnly = true;
            this.MsgWindow.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.MsgWindow.Size = new System.Drawing.Size(490, 208);
            this.MsgWindow.TabIndex = 1;
            // 
            // btnSelectBugFile
            // 
            this.btnSelectBugFile.Location = new System.Drawing.Point(435, 9);
            this.btnSelectBugFile.Name = "btnSelectBugFile";
            this.btnSelectBugFile.Size = new System.Drawing.Size(67, 23);
            this.btnSelectBugFile.TabIndex = 2;
            this.btnSelectBugFile.Text = "Select";
            this.btnSelectBugFile.UseVisualStyleBackColor = true;
            this.btnSelectBugFile.Click += new System.EventHandler(this.btnSelectBugFile_Click);
            // 
            // txtBugFile
            // 
            this.txtBugFile.Location = new System.Drawing.Point(99, 9);
            this.txtBugFile.Name = "txtBugFile";
            this.txtBugFile.ReadOnly = true;
            this.txtBugFile.Size = new System.Drawing.Size(330, 22);
            this.txtBugFile.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 12);
            this.label1.TabIndex = 4;
            this.label1.Text = "Bug File";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 43);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(69, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "Test Case File";
            // 
            // txtTCFile
            // 
            this.txtTCFile.Location = new System.Drawing.Point(99, 37);
            this.txtTCFile.Name = "txtTCFile";
            this.txtTCFile.ReadOnly = true;
            this.txtTCFile.Size = new System.Drawing.Size(330, 22);
            this.txtTCFile.TabIndex = 6;
            // 
            // btnSelectTCFile
            // 
            this.btnSelectTCFile.Location = new System.Drawing.Point(435, 36);
            this.btnSelectTCFile.Name = "btnSelectTCFile";
            this.btnSelectTCFile.Size = new System.Drawing.Size(67, 23);
            this.btnSelectTCFile.TabIndex = 7;
            this.btnSelectTCFile.Text = "Select";
            this.btnSelectTCFile.UseVisualStyleBackColor = true;
            this.btnSelectTCFile.Click += new System.EventHandler(this.btnSelectTCFile_Click);
            // 
            // btnSelectReportFile
            // 
            this.btnSelectReportFile.Location = new System.Drawing.Point(435, 64);
            this.btnSelectReportFile.Name = "btnSelectReportFile";
            this.btnSelectReportFile.Size = new System.Drawing.Size(67, 23);
            this.btnSelectReportFile.TabIndex = 10;
            this.btnSelectReportFile.Text = "Select";
            this.btnSelectReportFile.UseVisualStyleBackColor = true;
            this.btnSelectReportFile.Click += new System.EventHandler(this.btnSelectReportFile_Click);
            // 
            // txtReportFile
            // 
            this.txtReportFile.Location = new System.Drawing.Point(99, 65);
            this.txtReportFile.Name = "txtReportFile";
            this.txtReportFile.ReadOnly = true;
            this.txtReportFile.Size = new System.Drawing.Size(330, 22);
            this.txtReportFile.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 71);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(83, 12);
            this.label3.TabIndex = 8;
            this.label3.Text = "Report Template";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 128);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(24, 12);
            this.label4.TabIndex = 11;
            this.label4.Text = "Log";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(514, 372);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnSelectReportFile);
            this.Controls.Add(this.txtReportFile);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnSelectTCFile);
            this.Controls.Add(this.txtTCFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtBugFile);
            this.Controls.Add(this.btnSelectBugFile);
            this.Controls.Add(this.MsgWindow);
            this.Controls.Add(this.btn_run);
            this.Name = "MainForm";
            this.Text = "ReportGenerator";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_run;
        private System.Windows.Forms.TextBox MsgWindow;
        private System.Windows.Forms.Button btnSelectBugFile;
        private System.Windows.Forms.TextBox txtBugFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtTCFile;
        private System.Windows.Forms.Button btnSelectTCFile;
        private System.Windows.Forms.Button btnSelectReportFile;
        private System.Windows.Forms.TextBox txtReportFile;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}

