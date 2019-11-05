namespace ExcelAssessmentIntegration
{
    partial class ExcelIntegrationAssessmentWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExcelIntegrationAssessmentWindow));
            this.filterCriteriaGrpBx = new System.Windows.Forms.GroupBox();
            this.sectionRB = new System.Windows.Forms.RadioButton();
            this.courseRB = new System.Windows.Forms.RadioButton();
            this.semesterRB = new System.Windows.Forms.RadioButton();
            this.yearRB = new System.Windows.Forms.RadioButton();
            this.yearTBx = new System.Windows.Forms.TextBox();
            this.semesterCmBx = new System.Windows.Forms.ComboBox();
            this.courseCmBx = new System.Windows.Forms.ComboBox();
            this.sectionCmBx = new System.Windows.Forms.ComboBox();
            this.readExcelBtn = new System.Windows.Forms.Button();
            this.cancelBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.filterCriteriaGrpBx.SuspendLayout();
            this.SuspendLayout();
            // 
            // filterCriteriaGrpBx
            // 
            this.filterCriteriaGrpBx.Controls.Add(this.sectionRB);
            this.filterCriteriaGrpBx.Controls.Add(this.courseRB);
            this.filterCriteriaGrpBx.Controls.Add(this.semesterRB);
            this.filterCriteriaGrpBx.Controls.Add(this.yearRB);
            this.filterCriteriaGrpBx.Location = new System.Drawing.Point(18, 18);
            this.filterCriteriaGrpBx.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.filterCriteriaGrpBx.Name = "filterCriteriaGrpBx";
            this.filterCriteriaGrpBx.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.filterCriteriaGrpBx.Size = new System.Drawing.Size(159, 203);
            this.filterCriteriaGrpBx.TabIndex = 1;
            this.filterCriteriaGrpBx.TabStop = false;
            this.filterCriteriaGrpBx.Text = "Filter Criteria";
            // 
            // sectionRB
            // 
            this.sectionRB.AutoSize = true;
            this.sectionRB.Location = new System.Drawing.Point(10, 142);
            this.sectionRB.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.sectionRB.Name = "sectionRB";
            this.sectionRB.Size = new System.Drawing.Size(88, 24);
            this.sectionRB.TabIndex = 3;
            this.sectionRB.TabStop = true;
            this.sectionRB.Text = "Section";
            this.sectionRB.UseVisualStyleBackColor = true;
            this.sectionRB.CheckedChanged += new System.EventHandler(this.filterCriteriaGrpBx_CheckedChanged);
            // 
            // courseRB
            // 
            this.courseRB.AutoSize = true;
            this.courseRB.Location = new System.Drawing.Point(10, 105);
            this.courseRB.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.courseRB.Name = "courseRB";
            this.courseRB.Size = new System.Drawing.Size(85, 24);
            this.courseRB.TabIndex = 2;
            this.courseRB.TabStop = true;
            this.courseRB.Text = "Course";
            this.courseRB.UseVisualStyleBackColor = true;
            this.courseRB.CheckedChanged += new System.EventHandler(this.filterCriteriaGrpBx_CheckedChanged);
            // 
            // semesterRB
            // 
            this.semesterRB.AutoSize = true;
            this.semesterRB.Location = new System.Drawing.Point(10, 68);
            this.semesterRB.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.semesterRB.Name = "semesterRB";
            this.semesterRB.Size = new System.Drawing.Size(103, 24);
            this.semesterRB.TabIndex = 1;
            this.semesterRB.TabStop = true;
            this.semesterRB.Text = "Semester";
            this.semesterRB.UseVisualStyleBackColor = true;
            this.semesterRB.CheckedChanged += new System.EventHandler(this.filterCriteriaGrpBx_CheckedChanged);
            // 
            // yearRB
            // 
            this.yearRB.AutoSize = true;
            this.yearRB.Location = new System.Drawing.Point(10, 31);
            this.yearRB.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.yearRB.Name = "yearRB";
            this.yearRB.Size = new System.Drawing.Size(68, 24);
            this.yearRB.TabIndex = 0;
            this.yearRB.TabStop = true;
            this.yearRB.Text = "Year";
            this.yearRB.UseVisualStyleBackColor = true;
            this.yearRB.CheckedChanged += new System.EventHandler(this.filterCriteriaGrpBx_CheckedChanged);
            // 
            // yearTBx
            // 
            this.yearTBx.Location = new System.Drawing.Point(204, 45);
            this.yearTBx.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.yearTBx.MaxLength = 2;
            this.yearTBx.Name = "yearTBx";
            this.yearTBx.Size = new System.Drawing.Size(148, 26);
            this.yearTBx.TabIndex = 2;
            // 
            // semesterCmBx
            // 
            this.semesterCmBx.FormattingEnabled = true;
            this.semesterCmBx.Items.AddRange(new object[] {
            "Fall",
            "Spring",
            "Summer",
            "Winter"});
            this.semesterCmBx.Location = new System.Drawing.Point(402, 43);
            this.semesterCmBx.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.semesterCmBx.Name = "semesterCmBx";
            this.semesterCmBx.Size = new System.Drawing.Size(180, 28);
            this.semesterCmBx.TabIndex = 3;
            this.semesterCmBx.SelectedIndexChanged += new System.EventHandler(this.semesterCmBx_SelectedIndexChanged);
            // 
            // courseCmBx
            // 
            this.courseCmBx.FormattingEnabled = true;
            this.courseCmBx.Items.AddRange(new object[] {
            "CPSC 130",
            "CPSC 146",
            "CPSC 207",
            "CPSC 217",
            "CPSC 246",
            "CPSC 300",
            "CPSC 311",
            "CPSC 323",
            "CPSC 327",
            "CPSC 337",
            "CPSC 376",
            "CPSC 423",
            "CPSC 427",
            "CPSC 488"});
            this.courseCmBx.Location = new System.Drawing.Point(627, 43);
            this.courseCmBx.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.courseCmBx.Name = "courseCmBx";
            this.courseCmBx.Size = new System.Drawing.Size(180, 28);
            this.courseCmBx.TabIndex = 4;
            // 
            // sectionCmBx
            // 
            this.sectionCmBx.FormattingEnabled = true;
            this.sectionCmBx.Location = new System.Drawing.Point(849, 43);
            this.sectionCmBx.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.sectionCmBx.Name = "sectionCmBx";
            this.sectionCmBx.Size = new System.Drawing.Size(180, 28);
            this.sectionCmBx.TabIndex = 5;
            // 
            // readExcelBtn
            // 
            this.readExcelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.readExcelBtn.Location = new System.Drawing.Point(627, 151);
            this.readExcelBtn.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.readExcelBtn.Name = "readExcelBtn";
            this.readExcelBtn.Size = new System.Drawing.Size(204, 35);
            this.readExcelBtn.TabIndex = 6;
            this.readExcelBtn.Text = "Read Excel Sheets";
            this.readExcelBtn.UseVisualStyleBackColor = true;
            this.readExcelBtn.Click += new System.EventHandler(this.ReadExcelBtn_Click);
            // 
            // cancelBtn
            // 
            this.cancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelBtn.Location = new System.Drawing.Point(375, 151);
            this.cancelBtn.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cancelBtn.Name = "cancelBtn";
            this.cancelBtn.Size = new System.Drawing.Size(182, 35);
            this.cancelBtn.TabIndex = 7;
            this.cancelBtn.Text = "Cancel";
            this.cancelBtn.UseVisualStyleBackColor = true;
            this.cancelBtn.Click += new System.EventHandler(this.CancelBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(204, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(138, 20);
            this.label1.TabIndex = 8;
            this.label1.Text = "Year (Format: YY)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(402, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(78, 20);
            this.label2.TabIndex = 9;
            this.label2.Text = "Semester";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(627, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(60, 20);
            this.label3.TabIndex = 10;
            this.label3.Text = "Course";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(849, 13);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(63, 20);
            this.label4.TabIndex = 11;
            this.label4.Text = "Section";
            // 
            // ExcelIntegrationAssessmentWindow
            // 
            this.AcceptButton = this.readExcelBtn;
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelBtn;
            this.ClientSize = new System.Drawing.Size(1096, 323);
            this.ControlBox = false;
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cancelBtn);
            this.Controls.Add(this.readExcelBtn);
            this.Controls.Add(this.sectionCmBx);
            this.Controls.Add(this.courseCmBx);
            this.Controls.Add(this.semesterCmBx);
            this.Controls.Add(this.yearTBx);
            this.Controls.Add(this.filterCriteriaGrpBx);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "ExcelIntegrationAssessmentWindow";
            this.Text = "Excel Integration Assessment";
            this.filterCriteriaGrpBx.ResumeLayout(false);
            this.filterCriteriaGrpBx.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox filterCriteriaGrpBx;
        private System.Windows.Forms.RadioButton sectionRB;
        private System.Windows.Forms.RadioButton courseRB;
        private System.Windows.Forms.RadioButton semesterRB;
        private System.Windows.Forms.RadioButton yearRB;
        private System.Windows.Forms.TextBox yearTBx;
        private System.Windows.Forms.ComboBox semesterCmBx;
        private System.Windows.Forms.ComboBox courseCmBx;
        private System.Windows.Forms.ComboBox sectionCmBx;
        private System.Windows.Forms.Button readExcelBtn;
        private System.Windows.Forms.Button cancelBtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}

