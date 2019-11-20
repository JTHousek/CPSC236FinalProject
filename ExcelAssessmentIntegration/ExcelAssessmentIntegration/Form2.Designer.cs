namespace ExcelAssessmentIntegration
{
    partial class ProcessedWindow
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
            this.okBtn = new System.Windows.Forms.Button();
            this.outputObjLBx = new System.Windows.Forms.ListBox();
            this.outputPercLBx = new System.Windows.Forms.ListBox();
            this.outputNumLBx = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // okBtn
            // 
            this.okBtn.Location = new System.Drawing.Point(271, 303);
            this.okBtn.Name = "okBtn";
            this.okBtn.Size = new System.Drawing.Size(140, 32);
            this.okBtn.TabIndex = 0;
            this.okBtn.Text = "OK";
            this.okBtn.UseVisualStyleBackColor = true;
            this.okBtn.Click += new System.EventHandler(this.OkBtn_Click);
            // 
            // outputObjLBx
            // 
            this.outputObjLBx.FormattingEnabled = true;
            this.outputObjLBx.Location = new System.Drawing.Point(12, 23);
            this.outputObjLBx.Name = "outputObjLBx";
            this.outputObjLBx.Size = new System.Drawing.Size(386, 251);
            this.outputObjLBx.TabIndex = 1;
            // 
            // outputPercLBx
            // 
            this.outputPercLBx.FormattingEnabled = true;
            this.outputPercLBx.Location = new System.Drawing.Point(426, 23);
            this.outputPercLBx.Name = "outputPercLBx";
            this.outputPercLBx.Size = new System.Drawing.Size(95, 251);
            this.outputPercLBx.TabIndex = 2;
            // 
            // outputNumLBx
            // 
            this.outputNumLBx.FormattingEnabled = true;
            this.outputNumLBx.Location = new System.Drawing.Point(549, 23);
            this.outputNumLBx.Name = "outputNumLBx";
            this.outputNumLBx.Size = new System.Drawing.Size(95, 251);
            this.outputNumLBx.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Objectives";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(426, 4);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Percentages";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(549, 4);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "# of Students";
            // 
            // ProcessedWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(656, 365);
            this.ControlBox = false;
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.outputNumLBx);
            this.Controls.Add(this.outputPercLBx);
            this.Controls.Add(this.outputObjLBx);
            this.Controls.Add(this.okBtn);
            this.Name = "ProcessedWindow";
            this.ShowIcon = false;
            this.Text = "Assessment Comparisons";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button okBtn;
        private System.Windows.Forms.ListBox outputObjLBx;
        private System.Windows.Forms.ListBox outputPercLBx;
        private System.Windows.Forms.ListBox outputNumLBx;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
    }
}