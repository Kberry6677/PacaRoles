namespace PacaRoles
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
            this.btnExit = new System.Windows.Forms.Button();
            this.btnRun = new System.Windows.Forms.Button();
            this.txtFilename = new System.Windows.Forms.TextBox();
            this.btnGetFile = new System.Windows.Forms.Button();
            this.dlgFilePick = new System.Windows.Forms.OpenFileDialog();
            this.txtResults = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnExport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(197, 37);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(159, 23);
            this.btnExit.TabIndex = 8;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(16, 37);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(159, 23);
            this.btnRun.TabIndex = 7;
            this.btnRun.Text = "Run";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // txtFilename
            // 
            this.txtFilename.Location = new System.Drawing.Point(97, 11);
            this.txtFilename.Name = "txtFilename";
            this.txtFilename.ReadOnly = true;
            this.txtFilename.Size = new System.Drawing.Size(259, 20);
            this.txtFilename.TabIndex = 6;
            // 
            // btnGetFile
            // 
            this.btnGetFile.Location = new System.Drawing.Point(16, 8);
            this.btnGetFile.Name = "btnGetFile";
            this.btnGetFile.Size = new System.Drawing.Size(75, 23);
            this.btnGetFile.TabIndex = 5;
            this.btnGetFile.Text = "Get File";
            this.btnGetFile.UseVisualStyleBackColor = true;
            this.btnGetFile.Click += new System.EventHandler(this.btnGetFile_Click);
            // 
            // txtResults
            // 
            this.txtResults.Location = new System.Drawing.Point(16, 66);
            this.txtResults.Multiline = true;
            this.txtResults.Name = "txtResults";
            this.txtResults.ReadOnly = true;
            this.txtResults.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtResults.Size = new System.Drawing.Size(340, 289);
            this.txtResults.TabIndex = 10;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(197, 361);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(158, 23);
            this.btnExport.TabIndex = 11;
            this.btnExport.Text = "Export to Word";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(367, 392);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.txtResults);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.txtFilename);
            this.Controls.Add(this.btnGetFile);
            this.Name = "Form1";
            this.Text = "Get PACA Roles";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.TextBox txtFilename;
        private System.Windows.Forms.Button btnGetFile;
        private System.Windows.Forms.OpenFileDialog dlgFilePick;
        private System.Windows.Forms.TextBox txtResults;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnExport;
    }
}

