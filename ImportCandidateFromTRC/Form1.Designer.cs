namespace ImportCandidateFromTRC
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.txtFileLoc = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnImport = new System.Windows.Forms.Button();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.txtCurTerm = new System.Windows.Forms.TextBox();
            this.txtCurYear = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(303, 6);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(68, 20);
            this.btnBrowse.TabIndex = 0;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // txtFileLoc
            // 
            this.txtFileLoc.Location = new System.Drawing.Point(93, 6);
            this.txtFileLoc.Name = "txtFileLoc";
            this.txtFileLoc.ReadOnly = true;
            this.txtFileLoc.Size = new System.Drawing.Size(204, 20);
            this.txtFileLoc.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "CSV Location:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnImport
            // 
            this.btnImport.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImport.Location = new System.Drawing.Point(121, 98);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(140, 30);
            this.btnImport.TabIndex = 3;
            this.btnImport.Text = "Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // txtLog
            // 
            this.txtLog.Location = new System.Drawing.Point(15, 154);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLog.Size = new System.Drawing.Size(356, 105);
            this.txtLog.TabIndex = 4;
            // 
            // txtCurTerm
            // 
            this.txtCurTerm.Location = new System.Drawing.Point(211, 47);
            this.txtCurTerm.Name = "txtCurTerm";
            this.txtCurTerm.Size = new System.Drawing.Size(50, 20);
            this.txtCurTerm.TabIndex = 5;
            // 
            // txtCurYear
            // 
            this.txtCurYear.Location = new System.Drawing.Point(211, 73);
            this.txtCurYear.Name = "txtCurYear";
            this.txtCurYear.Size = new System.Drawing.Size(50, 20);
            this.txtCurYear.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(121, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 20);
            this.label2.TabIndex = 7;
            this.label2.Text = "Current Term:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(121, 73);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(84, 20);
            this.label3.TabIndex = 8;
            this.label3.Text = "Current Year:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(383, 267);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtCurYear);
            this.Controls.Add(this.txtCurTerm);
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtFileLoc);
            this.Controls.Add(this.btnBrowse);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Form1";
            this.Text = "Import Candidates from TRC";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtFileLoc;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.TextBox txtCurTerm;
        private System.Windows.Forms.TextBox txtCurYear;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
    }
}

