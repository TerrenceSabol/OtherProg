namespace DuplicateResolver
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
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnResolve = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(151, 5);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(49, 30);
            this.progressBar1.TabIndex = 13;
            // 
            // btnResolve
            // 
            this.btnResolve.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnResolve.Location = new System.Drawing.Point(97, 5);
            this.btnResolve.Name = "btnResolve";
            this.btnResolve.Size = new System.Drawing.Size(157, 30);
            this.btnResolve.TabIndex = 11;
            this.btnResolve.Text = "Resolve Duplicates";
            this.btnResolve.UseVisualStyleBackColor = true;
            this.btnResolve.Click += new System.EventHandler(this.btnResolve_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(336, 57);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnResolve);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Form1";
            this.Text = "Holding Duplicates Resolver";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        

        #endregion

        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btnResolve;
    }
}