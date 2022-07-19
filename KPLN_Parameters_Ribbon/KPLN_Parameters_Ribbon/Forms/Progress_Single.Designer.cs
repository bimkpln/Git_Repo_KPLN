namespace KPLN_Parameters_Ribbon.Forms
{
    partial class Progress_Single
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
            this.Header_lbl = new System.Windows.Forms.Label();
            this.Add_lbl = new System.Windows.Forms.Label();
            this.Titile_lbl = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 67);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(376, 30);
            this.progressBar1.TabIndex = 0;
            // 
            // Header_lbl
            // 
            this.Header_lbl.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this.Header_lbl.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Header_lbl.Location = new System.Drawing.Point(12, 9);
            this.Header_lbl.Name = "Header_lbl";
            this.Header_lbl.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Header_lbl.Size = new System.Drawing.Size(376, 55);
            this.Header_lbl.TabIndex = 1;
            this.Header_lbl.Text = "Инициализация...";
            this.Header_lbl.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.Header_lbl.UseWaitCursor = true;
            // 
            // Add_lbl
            // 
            this.Add_lbl.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Add_lbl.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.Add_lbl.Location = new System.Drawing.Point(13, 102);
            this.Add_lbl.Name = "Add_lbl";
            this.Add_lbl.Size = new System.Drawing.Size(375, 14);
            this.Add_lbl.TabIndex = 2;
            this.Add_lbl.Text = "...";
            // 
            // Titile_lbl
            // 
            this.Titile_lbl.AutoSize = true;
            this.Titile_lbl.Font = new System.Drawing.Font("Consolas", 9.75F);
            this.Titile_lbl.ForeColor = System.Drawing.Color.DarkGray;
            this.Titile_lbl.Location = new System.Drawing.Point(9, 9);
            this.Titile_lbl.Name = "Titile_lbl";
            this.Titile_lbl.Size = new System.Drawing.Size(98, 15);
            this.Titile_lbl.TabIndex = 3;
            this.Titile_lbl.Text = "Инициализация";
            // 
            // Progress_Single
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.HighlightText;
            this.ClientSize = new System.Drawing.Size(400, 125);
            this.ControlBox = false;
            this.Controls.Add(this.Titile_lbl);
            this.Controls.Add(this.Add_lbl);
            this.Controls.Add(this.Header_lbl);
            this.Controls.Add(this.progressBar1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(400, 125);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(400, 125);
            this.Name = "Progress_Single";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ProgressForm";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label Header_lbl;
        private System.Windows.Forms.Label Add_lbl;
        private System.Windows.Forms.Label Titile_lbl;
    }
}