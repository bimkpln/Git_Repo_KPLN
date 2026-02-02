using KPLN_DefaultPanelExtension_Modify.Forms.Models;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_DefaultPanelExtension_Modify.Forms
{
    public partial class ImgLargeFrom : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private readonly SendMsgToBitrix _mainView;
        private ImageSource _imageSource;
        private string _ilf_ImgSpecialFormat;

        public ImgLargeFrom(SendMsgToBitrix mainView, SendMsgToBitrix_VM mainViewModel)
        {
            _mainView = mainView;
            ILF_ViewModel = mainViewModel;

            InitializeComponent();

            Owner = mainView;
            DataContext = this;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public SendMsgToBitrix_VM ILF_ViewModel { get; set; }

        public ImageSource ILF_TaskImageSource
        {
            get
            {
                using (MemoryStream ms = new MemoryStream(ILF_ViewModel.ImageBuffer))
                {
                    BitmapImage bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = ms;
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();
                    _imageSource = bitmapImage;

                    return _imageSource;
                }
            }
            set
            {
                _imageSource = value;
                OnPropertyChanged();
            }
        }

        public string ILF_ImgSpecialFormat
        {
            get => _ilf_ImgSpecialFormat;
            set
            {
                _ilf_ImgSpecialFormat = value;
                OnPropertyChanged();
            }
        }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) Close();
        }

        private void ImgDelete_Click(object sender, RoutedEventArgs e) =>
            _mainView?.ImgDelete_Click(sender, e);

        private void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
