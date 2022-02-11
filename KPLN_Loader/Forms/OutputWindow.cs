using System.Drawing;
using System.Windows.Forms;

namespace KPLN_Loader.Forms
{
    public partial class OutputWindow : Form
    {
        public WebBrowser webBrowser = new WebBrowser();
        public OutputWindow()
        {
            InitializeComponent();
            this.SuspendLayout();
            this.webBrowser.Parent = this;
            this.webBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser.Location = new System.Drawing.Point(0, 0);
            this.webBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.Size = new System.Drawing.Size(800, 450);
            this.webBrowser.TabIndex = 0;
            this.webBrowser.Tag = "Web";
            this.webBrowser.ScrollBarsEnabled = true;
            this.webBrowser.AllowWebBrowserDrop = false;
            this.ResumeLayout(false);
            /*
            try
            {
                if (Preferences.RevitWindow.WindowState == System.Windows.WindowState.Normal)
                {
                    int x = (int)(Preferences.RevitWindow.Left) + (int)(Preferences.RevitWindow.Width) / 2 - 400;
                    int y = (int)(Preferences.RevitWindow.Top) + (int)(Preferences.RevitWindow.Height) / 2 - 300;
                    this.Location = new System.Drawing.Point(x, y);
                }
            }
            catch (System.Exception) { }
            */
        }

        private void OnClose(object sender, FormClosingEventArgs e)
        {
            Output.Output.FormOutput = null;
        }
    }
}
