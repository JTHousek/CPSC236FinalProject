namespace ExcelAssessmentIntegration
{
    partial class Form1
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
            this.filterCriteriaGrpBx.SuspendLayout();
            this.SuspendLayout();
            // 
            // filterCriteriaGrpBx
            // 
            this.filterCriteriaGrpBx.Controls.Add(this.sectionRB);
            this.filterCriteriaGrpBx.Controls.Add(this.courseRB);
            this.filterCriteriaGrpBx.Controls.Add(this.semesterRB);
            this.filterCriteriaGrpBx.Controls.Add(this.yearRB);
            this.filterCriteriaGrpBx.Location = new System.Drawing.Point(12, 12);
            this.filterCriteriaGrpBx.Name = "filterCriteriaGrpBx";
            this.filterCriteriaGrpBx.Size = new System.Drawing.Size(106, 132);
            this.filterCriteriaGrpBx.TabIndex = 1;
            this.filterCriteriaGrpBx.TabStop = false;
            this.filterCriteriaGrpBx.Text = "Filter Criteria";
            // 
            // sectionRB
            // 
            this.sectionRB.AutoSize = true;
            this.sectionRB.Location = new System.Drawing.Point(7, 92);
            this.sectionRB.Name = "sectionRB";
            this.sectionRB.Size = new System.Drawing.Size(61, 17);
            this.sectionRB.TabIndex = 3;
            this.sectionRB.TabStop = true;
            this.sectionRB.Text = "Section";
            this.sectionRB.UseVisualStyleBackColor = true;
            // 
            // courseRB
            // 
            this.courseRB.AutoSize = true;
            this.courseRB.Location = new System.Drawing.Point(7, 68);
            this.courseRB.Name = "courseRB";
            this.courseRB.Size = new System.Drawing.Size(58, 17);
            this.courseRB.TabIndex = 2;
            this.courseRB.TabStop = true;
            this.courseRB.Text = "Course";
            this.courseRB.UseVisualStyleBackColor = true;
            // 
            // semesterRB
            // 
            this.semesterRB.AutoSize = true;
            this.semesterRB.Location = new System.Drawing.Point(7, 44);
            this.semesterRB.Name = "semesterRB";
            this.semesterRB.Size = new System.Drawing.Size(69, 17);
            this.semesterRB.TabIndex = 1;
            this.semesterRB.TabStop = true;
            this.semesterRB.Text = "Semester";
            this.semesterRB.UseVisualStyleBackColor = true;
            // 
            // yearRB
            // 
            this.yearRB.AutoSize = true;
            this.yearRB.Location = new System.Drawing.Point(7, 20);
            this.yearRB.Name = "yearRB";
            this.yearRB.Size = new System.Drawing.Size(47, 17);
            this.yearRB.TabIndex = 0;
            this.yearRB.TabStop = true;
            this.yearRB.Text = "Year";
            this.yearRB.UseVisualStyleBackColor = true;
            // 
            // yearTBx
            // 
            this.yearTBx.Location = new System.Drawing.Point(136, 29);
            this.yearTBx.MaxLength = 2;
            this.yearTBx.Name = "yearTBx";
            this.yearTBx.Size = new System.Drawing.Size(100, 20);
            this.yearTBx.TabIndex = 2;
            // 
            // semesterCmBx
            // 
            this.semesterCmBx.FormattingEnabled = true;
            this.semesterCmBx.Items.AddRange(new object[] {
            "Fall",
            "Spring"});
            this.semesterCmBx.Location = new System.Drawing.Point(268, 28);
            this.semesterCmBx.Name = "semesterCmBx";
            this.semesterCmBx.Size = new System.Drawing.Size(121, 21);
            this.semesterCmBx.TabIndex = 3;
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
            this.courseCmBx.Location = new System.Drawing.Point(418, 28);
            this.courseCmBx.Name = "courseCmBx";
            this.courseCmBx.Size = new System.Drawing.Size(121, 21);
            this.courseCmBx.TabIndex = 4;
            // 
            // sectionCmBx
            // 
            this.sectionCmBx.FormattingEnabled = true;
            this.sectionCmBx.Location = new System.Drawing.Point(566, 28);
            this.sectionCmBx.Name = "sectionCmBx";
            this.sectionCmBx.Size = new System.Drawing.Size(121, 21);
            this.sectionCmBx.TabIndex = 5;
            // 
            // readExcelBtn
            // 
            this.readExcelBtn.Location = new System.Drawing.Point(338, 98);
            this.readExcelBtn.Name = "readExcelBtn";
            this.readExcelBtn.Size = new System.Drawing.Size(136, 23);
            this.readExcelBtn.TabIndex = 6;
            this.readExcelBtn.Text = "Read Excel Sheets";
            this.readExcelBtn.UseVisualStyleBackColor = true;
            this.readExcelBtn.Click += new System.EventHandler(this.ReadExcelBtn_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.readExcelBtn);
            this.Controls.Add(this.sectionCmBx);
            this.Controls.Add(this.courseCmBx);
            this.Controls.Add(this.semesterCmBx);
            this.Controls.Add(this.yearTBx);
            this.Controls.Add(this.filterCriteriaGrpBx);
            this.Name = "Form1";
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
    }
}

