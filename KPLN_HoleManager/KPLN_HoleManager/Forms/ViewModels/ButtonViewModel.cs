using System.Windows.Media.Imaging;

namespace KPLN_HoleManager.Forms.ViewModels
{
    internal class ButtonViewModel
    {
        public BitmapImage ImageSource { get; set; }
        public string MainText { get; set; }
        public string ToolTipText{ get; set; }

        public ButtonViewModel(BitmapImage imageSource, string mainText, string toolTipText)
        {
            ImageSource = imageSource;
            MainText = mainText;
            ToolTipText = toolTipText;
        }
    }
}
