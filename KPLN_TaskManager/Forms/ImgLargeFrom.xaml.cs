using KPLN_TaskManager.Common;
using System.Windows;
using System.Windows.Input;

namespace KPLN_TaskManager.Forms
{
    public partial class ImgLargeFrom : Window
    {
        public ImgLargeFrom(TaskItemEntity taskItemEntity)
        {
            InitializeComponent();

            DataContext = taskItemEntity;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) Close();
        }
    }
}
