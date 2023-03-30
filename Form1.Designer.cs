namespace SQLFinance
{
    partial class formMain
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
            this.lblStartDate = new System.Windows.Forms.Label();
            this.lblEndDate = new System.Windows.Forms.Label();
            this.cmdGo = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.cmdOFD = new System.Windows.Forms.Button();
            this.txtFileOutput = new System.Windows.Forms.TextBox();
            this.dtpStartDate = new System.Windows.Forms.DateTimePicker();
            this.dtpEndDate = new System.Windows.Forms.DateTimePicker();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            // 
            // lblStartDate
            // 
            this.lblStartDate.AutoSize = true;
            this.lblStartDate.Location = new System.Drawing.Point(12, 9);
            this.lblStartDate.Name = "lblStartDate";
            this.lblStartDate.Size = new System.Drawing.Size(58, 13);
            this.lblStartDate.TabIndex = 0;
            this.lblStartDate.Text = "Start Date:";
            // 
            // lblEndDate
            // 
            this.lblEndDate.AutoSize = true;
            this.lblEndDate.Location = new System.Drawing.Point(12, 49);
            this.lblEndDate.Name = "lblEndDate";
            this.lblEndDate.Size = new System.Drawing.Size(55, 13);
            this.lblEndDate.TabIndex = 1;
            this.lblEndDate.Text = "End Date:";
            // 
            // cmdGo
            // 
            this.cmdGo.Location = new System.Drawing.Point(294, 9);
            this.cmdGo.Name = "cmdGo";
            this.cmdGo.Size = new System.Drawing.Size(165, 52);
            this.cmdGo.TabIndex = 4;
            this.cmdGo.Text = "Go";
            this.cmdGo.UseVisualStyleBackColor = true;
            this.cmdGo.Click += new System.EventHandler(this.cmdGo_Click);
            // 
            // cmdOFD
            // 
            this.cmdOFD.Location = new System.Drawing.Point(12, 73);
            this.cmdOFD.Name = "cmdOFD";
            this.cmdOFD.Size = new System.Drawing.Size(58, 22);
            this.cmdOFD.TabIndex = 5;
            this.cmdOFD.Text = "Save As";
            this.cmdOFD.UseVisualStyleBackColor = true;
            this.cmdOFD.Click += new System.EventHandler(this.cmdOFD_Click);
            // 
            // txtFileOutput
            // 
            this.txtFileOutput.Enabled = false;
            this.txtFileOutput.Location = new System.Drawing.Point(76, 75);
            this.txtFileOutput.Name = "txtFileOutput";
            this.txtFileOutput.Size = new System.Drawing.Size(383, 20);
            this.txtFileOutput.TabIndex = 6;
            // 
            // dtpStartDate
            // 
            this.dtpStartDate.Location = new System.Drawing.Point(76, 9);
            this.dtpStartDate.MaxDate = new System.DateTime(2022, 5, 20, 0, 0, 0, 0);
            this.dtpStartDate.MinDate = new System.DateTime(2010, 1, 1, 0, 0, 0, 0);
            this.dtpStartDate.Name = "dtpStartDate";
            this.dtpStartDate.Size = new System.Drawing.Size(200, 20);
            this.dtpStartDate.TabIndex = 7;
            this.dtpStartDate.Value = new System.DateTime(2022, 5, 5, 0, 0, 0, 0);
            this.dtpStartDate.ValueChanged += new System.EventHandler(this.dtpStartDate_ValueChanged);
            // 
            // dtpEndDate
            // 
            this.dtpEndDate.Location = new System.Drawing.Point(76, 49);
            this.dtpEndDate.Name = "dtpEndDate";
            this.dtpEndDate.Size = new System.Drawing.Size(200, 20);
            this.dtpEndDate.TabIndex = 8;
            this.dtpEndDate.Value = new System.DateTime(2022, 6, 2, 10, 2, 33, 0);
            this.dtpEndDate.ValueChanged += new System.EventHandler(this.dtpEndDate_ValueChanged);
            // 
            // formMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(466, 105);
            this.Controls.Add(this.dtpEndDate);
            this.Controls.Add(this.dtpStartDate);
            this.Controls.Add(this.txtFileOutput);
            this.Controls.Add(this.cmdOFD);
            this.Controls.Add(this.cmdGo);
            this.Controls.Add(this.lblEndDate);
            this.Controls.Add(this.lblStartDate);
            this.Name = "formMain";
            this.Text = "Finance Report Generator";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.onClose);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void FormMain_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            throw new System.NotImplementedException();
        }

        #endregion

        private System.Windows.Forms.Label lblStartDate;
        private System.Windows.Forms.Label lblEndDate;
        private System.Windows.Forms.Button cmdGo;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button cmdOFD;
        private System.Windows.Forms.TextBox txtFileOutput;
        private System.Windows.Forms.DateTimePicker dtpStartDate;
        private System.Windows.Forms.DateTimePicker dtpEndDate;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}

