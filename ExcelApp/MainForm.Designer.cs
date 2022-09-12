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
            this.btnSelectBugFile = new System.Windows.Forms.Button();
            this.txtBugFile = new System.Windows.Forms.TextBox();
            this.label_issue = new System.Windows.Forms.Label();
            this.label_tc = new System.Windows.Forms.Label();
            this.txtTCFile = new System.Windows.Forms.TextBox();
            this.btnSelectTCFile = new System.Windows.Forms.Button();
            this.btnSelectReportFile = new System.Windows.Forms.Button();
            this.txtReportFile = new System.Windows.Forms.TextBox();
            this.label_1st = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnCreateReport = new System.Windows.Forms.Button();
            this.btnGetTCList = new System.Windows.Forms.Button();
            this.btnGetBugList = new System.Windows.Forms.Button();
            this.btnSelectExcelTestFile = new System.Windows.Forms.Button();
            this.txtExcelTestFile = new System.Windows.Forms.TextBox();
            this.label_2nd = new System.Windows.Forms.Label();
            this.comboBoxReportSelect = new System.Windows.Forms.ComboBox();
            this.MsgWindow = new System.Windows.Forms.TextBox();
            this.tabInfomation = new System.Windows.Forms.TabControl();
            this.tabReportInfo = new System.Windows.Forms.TabPage();
            this.txtReportInfo = new System.Windows.Forms.TextBox();
            this.tabExecutionLog = new System.Windows.Forms.TabPage();
            this.btnTestExcel = new System.Windows.Forms.Button();
            this.tabInfomation.SuspendLayout();
            this.tabReportInfo.SuspendLayout();
            this.tabExecutionLog.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnSelectBugFile
            // 
            this.btnSelectBugFile.Font = new System.Drawing.Font("Arial", 9F);
            this.btnSelectBugFile.Location = new System.Drawing.Point(418, 42);
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
            this.txtBugFile.Location = new System.Drawing.Point(110, 42);
            this.txtBugFile.Name = "txtBugFile";
            this.txtBugFile.ReadOnly = true;
            this.txtBugFile.Size = new System.Drawing.Size(302, 21);
            this.txtBugFile.TabIndex = 2;
            this.txtBugFile.TabStop = false;
            // 
            // label_issue
            // 
            this.label_issue.AutoSize = true;
            this.label_issue.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_issue.Location = new System.Drawing.Point(9, 47);
            this.label_issue.Name = "label_issue";
            this.label_issue.Size = new System.Drawing.Size(61, 15);
            this.label_issue.TabIndex = 4;
            this.label_issue.Text = "Issue File";
            // 
            // label_tc
            // 
            this.label_tc.AutoSize = true;
            this.label_tc.Font = new System.Drawing.Font("Arial", 9F);
            this.label_tc.Location = new System.Drawing.Point(9, 76);
            this.label_tc.Name = "label_tc";
            this.label_tc.Size = new System.Drawing.Size(86, 15);
            this.label_tc.TabIndex = 5;
            this.label_tc.Text = "Test Case File";
            // 
            // txtTCFile
            // 
            this.txtTCFile.Font = new System.Drawing.Font("Arial Narrow", 9F);
            this.txtTCFile.Location = new System.Drawing.Point(110, 70);
            this.txtTCFile.Name = "txtTCFile";
            this.txtTCFile.ReadOnly = true;
            this.txtTCFile.Size = new System.Drawing.Size(302, 21);
            this.txtTCFile.TabIndex = 6;
            this.txtTCFile.TabStop = false;
            // 
            // btnSelectTCFile
            // 
            this.btnSelectTCFile.Font = new System.Drawing.Font("Arial", 9F);
            this.btnSelectTCFile.Location = new System.Drawing.Point(418, 69);
            this.btnSelectTCFile.Name = "btnSelectTCFile";
            this.btnSelectTCFile.Size = new System.Drawing.Size(54, 23);
            this.btnSelectTCFile.TabIndex = 4;
            this.btnSelectTCFile.Text = "Select";
            this.btnSelectTCFile.UseVisualStyleBackColor = true;
            this.btnSelectTCFile.Click += new System.EventHandler(this.btnSelectTCFile_Click);
            // 
            // btnSelectReportFile
            // 
            this.btnSelectReportFile.Font = new System.Drawing.Font("Arial", 9F);
            this.btnSelectReportFile.Location = new System.Drawing.Point(418, 97);
            this.btnSelectReportFile.Name = "btnSelectReportFile";
            this.btnSelectReportFile.Size = new System.Drawing.Size(54, 23);
            this.btnSelectReportFile.TabIndex = 6;
            this.btnSelectReportFile.Text = "Select";
            this.btnSelectReportFile.UseVisualStyleBackColor = true;
            this.btnSelectReportFile.Click += new System.EventHandler(this.btnSelectReportFile_Click);
            // 
            // txtReportFile
            // 
            this.txtReportFile.Font = new System.Drawing.Font("Arial Narrow", 9F);
            this.txtReportFile.Location = new System.Drawing.Point(110, 99);
            this.txtReportFile.Name = "txtReportFile";
            this.txtReportFile.ReadOnly = true;
            this.txtReportFile.Size = new System.Drawing.Size(302, 21);
            this.txtReportFile.TabIndex = 9;
            this.txtReportFile.TabStop = false;
            // 
            // label_1st
            // 
            this.label_1st.AutoSize = true;
            this.label_1st.Font = new System.Drawing.Font("Arial", 9F);
            this.label_1st.Location = new System.Drawing.Point(9, 104);
            this.label_1st.Name = "label_1st";
            this.label_1st.Size = new System.Drawing.Size(44, 15);
            this.label_1st.TabIndex = 8;
            this.label_1st.Text = "Report";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(8, 15);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(77, 15);
            this.label4.TabIndex = 11;
            this.label4.Text = "Report Type ";
            // 
            // btnCreateReport
            // 
            this.btnCreateReport.Font = new System.Drawing.Font("Arial", 9F);
            this.btnCreateReport.Location = new System.Drawing.Point(478, 11);
            this.btnCreateReport.Name = "btnCreateReport";
            this.btnCreateReport.Size = new System.Drawing.Size(54, 25);
            this.btnCreateReport.TabIndex = 7;
            this.btnCreateReport.Text = "Report";
            this.btnCreateReport.UseVisualStyleBackColor = true;
            this.btnCreateReport.Click += new System.EventHandler(this.btnCreateReport_Click);
            // 
            // btnGetTCList
            // 
            this.btnGetTCList.Font = new System.Drawing.Font("Arial", 9F);
            this.btnGetTCList.Location = new System.Drawing.Point(477, 70);
            this.btnGetTCList.Name = "btnGetTCList";
            this.btnGetTCList.Size = new System.Drawing.Size(54, 23);
            this.btnGetTCList.TabIndex = 5;
            this.btnGetTCList.Text = "Read";
            this.btnGetTCList.UseVisualStyleBackColor = true;
            this.btnGetTCList.Click += new System.EventHandler(this.btnGetTCList_Click);
            // 
            // btnGetBugList
            // 
            this.btnGetBugList.Font = new System.Drawing.Font("Arial", 9F);
            this.btnGetBugList.Location = new System.Drawing.Point(478, 42);
            this.btnGetBugList.Name = "btnGetBugList";
            this.btnGetBugList.Size = new System.Drawing.Size(54, 23);
            this.btnGetBugList.TabIndex = 3;
            this.btnGetBugList.Text = "Read";
            this.btnGetBugList.UseVisualStyleBackColor = true;
            this.btnGetBugList.Click += new System.EventHandler(this.btnGetBugList_Click);
            // 
            // btnSelectExcelTestFile
            // 
            this.btnSelectExcelTestFile.Font = new System.Drawing.Font("Arial", 9F);
            this.btnSelectExcelTestFile.Location = new System.Drawing.Point(418, 124);
            this.btnSelectExcelTestFile.Name = "btnSelectExcelTestFile";
            this.btnSelectExcelTestFile.Size = new System.Drawing.Size(54, 23);
            this.btnSelectExcelTestFile.TabIndex = 8;
            this.btnSelectExcelTestFile.Text = "Select";
            this.btnSelectExcelTestFile.UseVisualStyleBackColor = true;
            this.btnSelectExcelTestFile.Click += new System.EventHandler(this.btnSelectExcelTestFile_Click);
            // 
            // txtExcelTestFile
            // 
            this.txtExcelTestFile.Font = new System.Drawing.Font("Arial Narrow", 9F);
            this.txtExcelTestFile.Location = new System.Drawing.Point(110, 126);
            this.txtExcelTestFile.Name = "txtExcelTestFile";
            this.txtExcelTestFile.ReadOnly = true;
            this.txtExcelTestFile.Size = new System.Drawing.Size(302, 21);
            this.txtExcelTestFile.TabIndex = 16;
            this.txtExcelTestFile.TabStop = false;
            // 
            // label_2nd
            // 
            this.label_2nd.AutoSize = true;
            this.label_2nd.Font = new System.Drawing.Font("Arial", 9F);
            this.label_2nd.Location = new System.Drawing.Point(9, 131);
            this.label_2nd.Name = "label_2nd";
            this.label_2nd.Size = new System.Drawing.Size(74, 15);
            this.label_2nd.TabIndex = 15;
            this.label_2nd.Text = "Extra Report";
            // 
            // comboBoxReportSelect
            // 
            this.comboBoxReportSelect.BackColor = System.Drawing.SystemColors.Control;
            this.comboBoxReportSelect.Font = new System.Drawing.Font("Arial Narrow", 9F);
            this.comboBoxReportSelect.FormattingEnabled = true;
            this.comboBoxReportSelect.Location = new System.Drawing.Point(110, 11);
            this.comboBoxReportSelect.Name = "comboBoxReportSelect";
            this.comboBoxReportSelect.Size = new System.Drawing.Size(302, 24);
            this.comboBoxReportSelect.TabIndex = 1;
            this.comboBoxReportSelect.SelectedIndexChanged += new System.EventHandler(this.comboBoxReportSelect_SelectedIndexChanged);
            // 
            // MsgWindow
            // 
            this.MsgWindow.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MsgWindow.Location = new System.Drawing.Point(-3, 0);
            this.MsgWindow.Multiline = true;
            this.MsgWindow.Name = "MsgWindow";
            this.MsgWindow.ReadOnly = true;
            this.MsgWindow.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.MsgWindow.Size = new System.Drawing.Size(516, 196);
            this.MsgWindow.TabIndex = 1;
            this.MsgWindow.TabStop = false;
            // 
            // tabInfomation
            // 
            this.tabInfomation.Controls.Add(this.tabReportInfo);
            this.tabInfomation.Controls.Add(this.tabExecutionLog);
            this.tabInfomation.Font = new System.Drawing.Font("Arial", 9F);
            this.tabInfomation.Location = new System.Drawing.Point(11, 157);
            this.tabInfomation.Name = "tabInfomation";
            this.tabInfomation.SelectedIndex = 0;
            this.tabInfomation.Size = new System.Drawing.Size(521, 220);
            this.tabInfomation.TabIndex = 22;
            // 
            // tabReportInfo
            // 
            this.tabReportInfo.BackColor = System.Drawing.SystemColors.Control;
            this.tabReportInfo.Controls.Add(this.txtReportInfo);
            this.tabReportInfo.ForeColor = System.Drawing.SystemColors.ControlText;
            this.tabReportInfo.Location = new System.Drawing.Point(4, 24);
            this.tabReportInfo.Name = "tabReportInfo";
            this.tabReportInfo.Padding = new System.Windows.Forms.Padding(3);
            this.tabReportInfo.Size = new System.Drawing.Size(513, 192);
            this.tabReportInfo.TabIndex = 0;
            this.tabReportInfo.Text = "Info";
            // 
            // txtReportInfo
            // 
            this.txtReportInfo.BackColor = System.Drawing.SystemColors.Control;
            this.txtReportInfo.Location = new System.Drawing.Point(-4, 0);
            this.txtReportInfo.Multiline = true;
            this.txtReportInfo.Name = "txtReportInfo";
            this.txtReportInfo.ReadOnly = true;
            this.txtReportInfo.Size = new System.Drawing.Size(520, 192);
            this.txtReportInfo.TabIndex = 2;
            this.txtReportInfo.TabStop = false;
            // 
            // tabExecutionLog
            // 
            this.tabExecutionLog.Controls.Add(this.MsgWindow);
            this.tabExecutionLog.Location = new System.Drawing.Point(4, 24);
            this.tabExecutionLog.Name = "tabExecutionLog";
            this.tabExecutionLog.Padding = new System.Windows.Forms.Padding(3);
            this.tabExecutionLog.Size = new System.Drawing.Size(513, 192);
            this.tabExecutionLog.TabIndex = 1;
            this.tabExecutionLog.Text = "Log";
            this.tabExecutionLog.UseVisualStyleBackColor = true;
            // 
            // btnTestExcel
            // 
            this.btnTestExcel.Font = new System.Drawing.Font("Arial", 9F);
            this.btnTestExcel.Location = new System.Drawing.Point(478, 124);
            this.btnTestExcel.Name = "btnTestExcel";
            this.btnTestExcel.Size = new System.Drawing.Size(54, 23);
            this.btnTestExcel.TabIndex = 9;
            this.btnTestExcel.Text = "Test";
            this.btnTestExcel.UseVisualStyleBackColor = true;
            this.btnTestExcel.Click += new System.EventHandler(this.btnTestExcel_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(544, 381);
            this.Controls.Add(this.tabInfomation);
            this.Controls.Add(this.comboBoxReportSelect);
            this.Controls.Add(this.btnTestExcel);
            this.Controls.Add(this.btnSelectExcelTestFile);
            this.Controls.Add(this.txtExcelTestFile);
            this.Controls.Add(this.label_2nd);
            this.Controls.Add(this.btnCreateReport);
            this.Controls.Add(this.btnGetTCList);
            this.Controls.Add(this.btnGetBugList);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnSelectReportFile);
            this.Controls.Add(this.txtReportFile);
            this.Controls.Add(this.label_1st);
            this.Controls.Add(this.btnSelectTCFile);
            this.Controls.Add(this.txtTCFile);
            this.Controls.Add(this.label_tc);
            this.Controls.Add(this.label_issue);
            this.Controls.Add(this.txtBugFile);
            this.Controls.Add(this.btnSelectBugFile);
            this.Name = "MainForm";
            this.Text = "ReportGenerator";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.tabInfomation.ResumeLayout(false);
            this.tabReportInfo.ResumeLayout(false);
            this.tabReportInfo.PerformLayout();
            this.tabExecutionLog.ResumeLayout(false);
            this.tabExecutionLog.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSelectBugFile;
        private System.Windows.Forms.TextBox txtBugFile;
        private System.Windows.Forms.Label label_issue;
        private System.Windows.Forms.Label label_tc;
        private System.Windows.Forms.TextBox txtTCFile;
        private System.Windows.Forms.Button btnSelectTCFile;
        private System.Windows.Forms.Button btnSelectReportFile;
        private System.Windows.Forms.TextBox txtReportFile;
        private System.Windows.Forms.Label label_1st;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnCreateReport;
        private System.Windows.Forms.Button btnGetTCList;
        private System.Windows.Forms.Button btnGetBugList;
        private System.Windows.Forms.Button btnSelectExcelTestFile;
        private System.Windows.Forms.TextBox txtExcelTestFile;
        private System.Windows.Forms.Label label_2nd;
        private System.Windows.Forms.ComboBox comboBoxReportSelect;
        private System.Windows.Forms.TextBox MsgWindow;
        private System.Windows.Forms.TabControl tabInfomation;
        private System.Windows.Forms.TabPage tabReportInfo;
        private System.Windows.Forms.TabPage tabExecutionLog;
        private System.Windows.Forms.Button btnTestExcel;
        private System.Windows.Forms.TextBox txtReportInfo;
    }
}

