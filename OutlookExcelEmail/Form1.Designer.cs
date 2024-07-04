namespace OutlookExcelEmail
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.btnTest = new System.Windows.Forms.Button();
            this.btnSend = new System.Windows.Forms.Button();
            this.txtExcelFilePath = new System.Windows.Forms.TextBox();
            this.txtTestEmail = new System.Windows.Forms.TextBox();
            this.txtSendEmail = new System.Windows.Forms.TextBox();
            this.txtEmailSubject = new System.Windows.Forms.TextBox();
            this.lblExcelFilePath = new System.Windows.Forms.Label();
            this.lblTestEmail = new System.Windows.Forms.Label();
            this.lblSendEmail = new System.Windows.Forms.Label();
            this.lblEmailSubject = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(15, 200);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(75, 23);
            this.btnTest.TabIndex = 0;
            this.btnTest.Text = "Test";
            this.btnTest.UseVisualStyleBackColor = true;
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(110, 200);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(75, 23);
            this.btnSend.TabIndex = 1;
            this.btnSend.Text = "Send";
            this.btnSend.UseVisualStyleBackColor = true;
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // txtExcelFilePath
            // 
            this.txtExcelFilePath.Location = new System.Drawing.Point(15, 25);
            this.txtExcelFilePath.Name = "txtExcelFilePath";
            this.txtExcelFilePath.Size = new System.Drawing.Size(255, 20);
            this.txtExcelFilePath.TabIndex = 2;
            this.txtExcelFilePath.Text = "";
            // 
            // txtTestEmail
            // 
            this.txtTestEmail.Location = new System.Drawing.Point(15, 75);
            this.txtTestEmail.Name = "txtTestEmail";
            this.txtTestEmail.Size = new System.Drawing.Size(255, 20);
            this.txtTestEmail.TabIndex = 3;
            this.txtTestEmail.Text = "phw8484@hit12.co.kr";

            // 
            // txtSendEmail
            // 
            this.txtSendEmail.Location = new System.Drawing.Point(15, 125);
            this.txtSendEmail.Name = "txtSendEmail";
            this.txtSendEmail.Size = new System.Drawing.Size(255, 20);
            this.txtSendEmail.TabIndex = 4;
            
            // 
            // txtEmailSubject
            // 
            this.txtEmailSubject.Location = new System.Drawing.Point(15, 175);
            this.txtEmailSubject.Name = "txtEmailSubject";
            this.txtEmailSubject.Size = new System.Drawing.Size(255, 20);
            this.txtEmailSubject.TabIndex = 5;
            
            // 
            // lblExcelFilePath
            // 
            this.lblExcelFilePath.AutoSize = true;
            this.lblExcelFilePath.Location = new System.Drawing.Point(12, 9);
            this.lblExcelFilePath.Name = "lblExcelFilePath";
            this.lblExcelFilePath.Size = new System.Drawing.Size(80, 13);
            this.lblExcelFilePath.TabIndex = 6;
            this.lblExcelFilePath.Text = "Excel File Path:";
            // 
            // lblTestEmail
            // 
            this.lblTestEmail.AutoSize = true;
            this.lblTestEmail.Location = new System.Drawing.Point(12, 59);
            this.lblTestEmail.Name = "lblTestEmail";
            this.lblTestEmail.Size = new System.Drawing.Size(62, 13);
            this.lblTestEmail.TabIndex = 7;
            this.lblTestEmail.Text = "Test Email:";
            // 
            // lblSendEmail
            // 
            this.lblSendEmail.AutoSize = true;
            this.lblSendEmail.Location = new System.Drawing.Point(12, 109);
            this.lblSendEmail.Name = "lblSendEmail";
            this.lblSendEmail.Size = new System.Drawing.Size(64, 13);
            this.lblSendEmail.TabIndex = 8;
            this.lblSendEmail.Text = "Send Email:";
            // 
            // lblEmailSubject
            // 
            this.lblEmailSubject.AutoSize = true;
            this.lblEmailSubject.Location = new System.Drawing.Point(12, 159);
            this.lblEmailSubject.Name = "lblEmailSubject";
            this.lblEmailSubject.Size = new System.Drawing.Size(75, 13);
            this.lblEmailSubject.TabIndex = 9;
            this.lblEmailSubject.Text = "Email Subject:";
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(284, 241);
            this.Controls.Add(this.lblEmailSubject);
            this.Controls.Add(this.lblSendEmail);
            this.Controls.Add(this.lblTestEmail);
            this.Controls.Add(this.lblExcelFilePath);
            this.Controls.Add(this.txtEmailSubject);
            this.Controls.Add(this.txtSendEmail);
            this.Controls.Add(this.txtTestEmail);
            this.Controls.Add(this.txtExcelFilePath);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.btnTest);
            this.Name = "Form1";
            this.Text = "Outlook Email Sender";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.TextBox txtExcelFilePath;
        private System.Windows.Forms.TextBox txtTestEmail;
        private System.Windows.Forms.TextBox txtSendEmail;
        private System.Windows.Forms.TextBox txtEmailSubject;
        private System.Windows.Forms.Label lblExcelFilePath;
        private System.Windows.Forms.Label lblSendEmail;
        private System.Windows.Forms.Label lblEmailSubject;
        private System.Windows.Forms.Label lblTestEmail;

    }
}
