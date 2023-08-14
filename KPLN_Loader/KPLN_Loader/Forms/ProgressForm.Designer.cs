using System.Drawing;
using System.Reflection;

namespace KPLN_Loader.Forms
{
    partial class ProgressForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer _components = null;
        private System.Windows.Forms.Label _title;
        private System.Windows.Forms.Label _header;
        private System.Windows.Forms.ProgressBar _progressBar;
        private System.Windows.Forms.Label _description;
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (_components != null))
            {
                _components.Dispose();
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
            this._title = new System.Windows.Forms.Label();
            this._header = new System.Windows.Forms.Label();
            this._progressBar = new System.Windows.Forms.ProgressBar();
            this._description = new System.Windows.Forms.Label();
            this._loaderVersion = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // _title
            // 
            this._title.Location = new System.Drawing.Point(0, 0);
            this._title.Name = "_title";
            this._title.Size = new System.Drawing.Size(100, 50);
            this._title.TabIndex = 0;
            // 
            // _header
            // 
            this._header.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this._header.Font = new System.Drawing.Font("GOST Common", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._header.Location = new System.Drawing.Point(5, 15);
            this._header.Margin = new System.Windows.Forms.Padding(0);
            this._header.Name = "_header";
            this._header.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this._header.Size = new System.Drawing.Size(360, 27);
            this._header.TabIndex = 1;
            this._header.Text = "Заголовок указывается через констурктор";
            this._header.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this._header.UseWaitCursor = true;
            // 
            // _progressBar
            // 
            this._progressBar.ForeColor = System.Drawing.Color.Red;
            this._progressBar.Location = new System.Drawing.Point(5, 45);
            this._progressBar.Margin = new System.Windows.Forms.Padding(5, 0, 0, 5);
            this._progressBar.Name = "_progressBar";
            this._progressBar.Size = new System.Drawing.Size(360, 20);
            this._progressBar.Step = 1;
            this._progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this._progressBar.TabIndex = 0;
            // 
            // _description
            // 
            this._description.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._description.AutoEllipsis = true;
            this._description.AutoSize = true;
            this._description.Font = new System.Drawing.Font("GOST Common", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._description.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this._description.Location = new System.Drawing.Point(5, 70);
            this._description.Margin = new System.Windows.Forms.Padding(0, 5, 0, 5);
            this._description.MaximumSize = new System.Drawing.Size(385, 200);
            this._description.Name = "_description";
            this._description.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this._description.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this._description.Size = new System.Drawing.Size(310, 19);
            this._description.TabIndex = 2;
            this._description.Text = "Описание указывается через конструктор";
            // 
            // _loaderVersion
            // 
            this._loaderVersion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this._loaderVersion.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this._loaderVersion.Font = new System.Drawing.Font("GOST Common", 8F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this._loaderVersion.ForeColor = System.Drawing.SystemColors.ControlDark;
            this._loaderVersion.Location = new System.Drawing.Point(0, 0);
            this._loaderVersion.Margin = new System.Windows.Forms.Padding(0);
            this._loaderVersion.Name = "_loaderVersion";
            this._loaderVersion.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this._loaderVersion.Size = new System.Drawing.Size(360, 15);
            this._loaderVersion.TabIndex = 3;
            this._loaderVersion.Text = "v." + Assembly.GetExecutingAssembly().GetName().Version.ToString();
            this._loaderVersion.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this._loaderVersion.UseWaitCursor = true;
            // 
            // ProgressForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoScroll = true;
            this.BackColor = System.Drawing.SystemColors.HighlightText;
            this.ClientSize = new System.Drawing.Size(369, 90);
            this.Controls.Add(this._loaderVersion);
            this.Controls.Add(this._description);
            this.Controls.Add(this._header);
            this.Controls.Add(this._progressBar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(385, 600);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(385, 100);
            this.Name = "ProgressForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Загрузка модулей KPLN";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label _loaderVersion;
    }
}