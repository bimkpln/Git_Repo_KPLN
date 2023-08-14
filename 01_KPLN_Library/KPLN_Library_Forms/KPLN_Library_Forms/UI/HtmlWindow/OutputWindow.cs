using System.Drawing;
using System.Windows.Forms;

namespace KPLN_Library_Forms.UI.HtmlWindow
{
    public partial class OutputWindow : Form
    {
        public WebBrowser webBrowser = new WebBrowser();

        public OutputWindow()
        {
            InitializeComponent();
            this.SuspendLayout();
            this.webBrowser.Parent = this;
            this.webBrowser.Dock = DockStyle.Fill;
            this.webBrowser.Location = new Point(0, 0);
            this.webBrowser.MinimumSize = new Size(20, 20);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.Size = new Size(800, 450);
            this.webBrowser.TabIndex = 0;
            this.webBrowser.Tag = "Web";
            this.webBrowser.ScrollBarsEnabled = true;
            this.webBrowser.AllowWebBrowserDrop = false;
            this.ResumeLayout(false);
        }

        private void OnClose(object sender, FormClosingEventArgs e)
        {
            HtmlOutput.FormOutput = null;
        }
    }
}
