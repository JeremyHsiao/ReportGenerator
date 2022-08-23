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
            this.btnDemo = new System.Windows.Forms.Button();
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
            this.btnCreateReport = new System.Windows.Forms.Button();
            this.btnGetTCList = new System.Windows.Forms.Button();
            this.btnGetBugList = new System.Windows.Forms.Button();
            this.btnTestExcel = new System.Windows.Forms.Button();
            this.btnSelectExcelTestFile = new System.Windows.Forms.Button();
            this.txtExcelTestFile = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnDemo
            // 
            this.btnDemo.Font = new System.Drawing.Font("Arial", 9F);
            this.btnDemo.Location = new System.Drawing.Point(478, 123);
            this.btnDemo.Name = "btnDemo";
            this.btnDemo.Size = new System.Drawing.Size(54, 23);
            this.btnDemo.TabIndex = 0;
            this.btnDemo.Text = "Demo";
            this.btnDemo.UseVisualStyleBackColor = true;
            this.btnDemo.Click += new System.EventHandler(this.btnDemo_Click);
            // 
            // MsgWindow
            // 
            this.MsgWindow.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MsgWindow.Location = new System.Drawing.Point(12, 152);
            this.MsgWindow.Multiline = true;
            this.MsgWindow.Name = "MsgWindow";
            this.MsgWindow.ReadOnly = true;
            this.MsgWindow.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.MsgWindow.Size = new System.Drawing.Size(520, 217);
            this.MsgWindow.TabIndex = 1;
            // 
            // btnSelectBugFile
            // 
            this.btnSelectBugFile.Font = new System.Drawing.Font("Arial", 9F);
            this.btnSelectBugFile.Location = new System.Drawing.Point(419, 9);
            this.btnSelectBugFile.Name = "btnSelectBugFile";
            this.btnSelectBugFile.Size = new System.Drawing.Size(54, 23);
            this.btnSelectBugFile.TabIndex = 2;
            this.btnSelectBugFile.Text = "Select";
            this.btnSelectBugFile.UseVisualStyleBackColor = true;
            this.btnSelectBugFile.Click += new System.EventHandler(this.btnSelectBugFile_Click);
            // 
            // txtBugFile
            // 
            this.txtBugFile.Font = new System.Drawing.Font("Arial Narrow", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBugFile.Location = new System.Drawing.Point(111, 9);
            this.txtBugFile.Name = "txtBugFile";
            this.txtBugFile.ReadOnly = true;
            this.txtBugFile.Size = new System.Drawing.Size(302, 21);
            this.txtBugFile.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(10, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 15);
            this.label1.TabIndex = 4;
            this.label1.Text = "Issue File";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 9F);
            this.label2.Location = new System.Drawing.Point(10, 43);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(86, 15);
            this.label2.TabIndex = 5;
            this.label2.Text = "Test Case File";
            // 
            // txtTCFile
            // 
            this.txtTCFile.Font = new System.Drawing.Font("Arial Narrow", 9F);
            this.txtTCFile.Location = new System.Drawing.Point(111, 37);
            this.txtTCFile.Name = "txtTCFile";
            this.txtTCFile.ReadOnly = true;
            this.txtTCFile.Size = new System.Drawing.Size(302, 21);
            this.txtTCFile.TabIndex = 6;
            // 
            // btnSelectTCFile
            // 
            this.btnSelectTCFile.Font = new System.Drawing.Font("Arial", 9F);
            this.btnSelectTCFile.Location = new System.Drawing.Point(419, 36);
            this.btnSelectTCFile.Name = "btnSelectTCFile";
            this.btnSelectTCFile.Size = new System.Drawing.Size(54, 23);
            this.btnSelectTCFile.TabIndex = 7;
            this.btnSelectTCFile.Text = "Select";
            this.btnSelectTCFile.UseVisualStyleBackColor = true;
            this.btnSelectTCFile.Click += new System.EventHandler(this.btnSelectTCFile_Click);
            // 
            // btnSelectReportFile
            // 
            this.btnSelectReportFile.Font = new System.Drawing.Font("Arial", 9F);
            this.btnSelectReportFile.Location = new System.Drawing.Point(419, 64);
            this.btnSelectReportFile.Name = "btnSelectReportFile";
            this.btnSelectReportFile.Size = new System.Drawing.Size(54, 23);
            this.btnSelectReportFile.TabIndex = 10;
            this.btnSelectReportFile.Text = "Select";
            this.btnSelectReportFile.UseVisualStyleBackColor = true;
            this.btnSelectReportFile.Click += new System.EventHandler(this.btnSelectReportFile_Click);
            // 
            // txtReportFile
            // 
            this.txtReportFile.Font = new System.Drawing.Font("Arial Narrow", 9F);
            this.txtReportFile.Location = new System.Drawing.Point(111, 66);
            this.txtReportFile.Name = "txtReportFile";
            this.txtReportFile.ReadOnly = true;
            this.txtReportFile.Size = new System.Drawing.Size(302, 21);
            this.txtReportFile.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial", 9F);
            this.label3.Location = new System.Drawing.Point(10, 71);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(98, 15);
            this.label3.TabIndex = 8;
            this.label3.Text = "Report Template";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(12, 128);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(28, 15);
            this.label4.TabIndex = 11;
            this.label4.Text = "Log";
            // 
            // btnCreateReport
            // 
            this.btnCreateReport.Font = new System.Drawing.Font("Arial", 9F);
            this.btnCreateReport.Location = new System.Drawing.Point(479, 64);
            this.btnCreateReport.Name = "btnCreateReport";
            this.btnCreateReport.Size = new System.Drawing.Size(54, 23);
            this.btnCreateReport.TabIndex = 14;
            this.btnCreateReport.Text = "Report";
            this.btnCreateReport.UseVisualStyleBackColor = true;
            this.btnCreateReport.Click += new System.EventHandler(this.btnCreateReport_Click);
            // 
            // btnGetTCList
            // 
            this.btnGetTCList.Font = new System.Drawing.Font("Arial", 9F);
            this.btnGetTCList.Location = new System.Drawing.Point(478, 37);
            this.btnGetTCList.Name = "btnGetTCList";
            this.btnGetTCList.Size = new System.Drawing.Size(54, 23);
            this.btnGetTCList.TabIndex = 13;
            this.btnGetTCList.Text = "Read";
            this.btnGetTCList.UseVisualStyleBackColor = true;
            this.btnGetTCList.Click += new System.EventHandler(this.btnGetTCList_Click);
            // 
            // btnGetBugList
            // 
            this.btnGetBugList.Font = new System.Drawing.Font("Arial", 9F);
            this.btnGetBugList.Location = new System.Drawing.Point(479, 9);
            this.btnGetBugList.Name = "btnGetBugList";
            this.btnGetBugList.Size = new System.Drawing.Size(54, 23);
            this.btnGetBugList.TabIndex = 12;
            this.btnGetBugList.Text = "Read";
            this.btnGetBugList.UseVisualStyleBackColor = true;
            this.btnGetBugList.Click += new System.EventHandler(this.btnGetBugList_Click);
            // 
            // btnTestExcel
            // 
            this.btnTestExcel.Font = new System.Drawing.Font("Arial", 9F);
            this.btnTestExcel.Location = new System.Drawing.Point(479, 91);
            this.btnTestExcel.Name = "btnTestExcel";
            this.btnTestExcel.Size = new System.Drawing.Size(54, 23);
            this.btnTestExcel.TabIndex = 18;
            this.btnTestExcel.Text = "Test";
            this.btnTestExcel.UseVisualStyleBackColor = true;
            this.btnTestExcel.Click += new System.EventHandler(this.btnTestExcel_Click);
            // 
            // btnSelectExcelTestFile
            // 
            this.btnSelectExcelTestFile.Font = new System.Drawing.Font("Arial", 9F);
            this.btnSelectExcelTestFile.Location = new System.Drawing.Point(419, 91);
            this.btnSelectExcelTestFile.Name = "btnSelectExcelTestFile";
            this.btnSelectExcelTestFile.Size = new System.Drawing.Size(54, 23);
            this.btnSelectExcelTestFile.TabIndex = 17;
            this.btnSelectExcelTestFile.Text = "Select";
            this.btnSelectExcelTestFile.UseVisualStyleBackColor = true;
            this.btnSelectExcelTestFile.Click += new System.EventHandler(this.btnSelectExcelTestFile_Click);
            // 
            // txtExcelTestFile
            // 
            this.txtExcelTestFile.Font = new System.Drawing.Font("Arial Narrow", 9F);
            this.txtExcelTestFile.Location = new System.Drawing.Point(111, 93);
            this.txtExcelTestFile.Name = "txtExcelTestFile";
            this.txtExcelTestFile.ReadOnly = true;
            this.txtExcelTestFile.Size = new System.Drawing.Size(302, 21);
            this.txtExcelTestFile.TabIndex = 16;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial", 9F);
            this.label5.Location = new System.Drawing.Point(10, 98);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(85, 15);
            this.label5.TabIndex = 15;
            this.label5.Text = "Excel Test File";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(544, 381);
            this.Controls.Add(this.btnTestExcel);
            this.Controls.Add(this.btnSelectExcelTestFile);
            this.Controls.Add(this.txtExcelTestFile);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnCreateReport);
            this.Controls.Add(this.btnGetTCList);
            this.Controls.Add(this.btnGetBugList);
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
            this.Controls.Add(this.btnDemo);
            this.Name = "MainForm";
            this.Text = "ReportGenerator";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnDemo;
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
        private System.Windows.Forms.Button btnCreateReport;
        private System.Windows.Forms.Button btnGetTCList;
        private System.Windows.Forms.Button btnGetBugList;
        private System.Windows.Forms.Button btnTestExcel;
        private System.Windows.Forms.Button btnSelectExcelTestFile;
        private System.Windows.Forms.TextBox txtExcelTestFile;
        private System.Windows.Forms.Label label5;
    }
}

