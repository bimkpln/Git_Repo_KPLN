using KPLN_TaskManager.Common;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_TaskManager.Forms
{
    public partial class ImgLargeFrom : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        
        private readonly TaskItemView _tiView;
        private ImageSource _imageSource;
        private string _ilf_ImgSpecialFormat;

        public ImgLargeFrom(TaskItemView tiView, TaskItemEntity tiEntity)
        {
            _tiView = tiView;
            TIEntity = tiEntity;
            ILF_TE_ImageBufferColl = tiEntity.TE_ImageBufferColl;
            ILF_ImgSpecialFormat = SetSpecialForm();

            InitializeComponent();

            DataContext = this;
            ILF_DeleteBtn.IsEnabled = _tiView.IsCreatorEditable;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public TaskItemEntity TIEntity { get; set; }
        
        public List<TaskEntity_ImageBuffer> ILF_TE_ImageBufferColl { get; set; }

        public ImageSource ILF_TaskImageSource
        {
            get
            {
                if (ILF_TE_ImageBufferColl == null || ILF_TE_ImageBufferColl.All(buff => buff.ImageBuffer == null || buff.ImageBuffer.Length == 0))
                    return null;

                using (MemoryStream ms = new MemoryStream(ILF_TE_ImageBufferColl[TIEntity.TE_ImageBuffer_Current].ImageBuffer))
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

        private void ImgLeft_Click(object sender, RoutedEventArgs e) 
        { 
            _tiView?.ImgLeft_Click(sender, e);
            ILF_ImgSpecialFormat = SetSpecialForm();
        }

        private void ImgRight_Click(object sender, RoutedEventArgs e)
        { 
            _tiView?.ImgRight_Click(sender, e);
            ILF_ImgSpecialFormat = SetSpecialForm();
        }

        private void ImgDelete_Click(object sender, RoutedEventArgs e)
        { 
            _tiView?.ImgDelete_Click(sender, e);
            ILF_ImgSpecialFormat = SetSpecialForm();
        }
        
        private void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        private string SetSpecialForm() => $"{TIEntity.TE_ImageBuffer_Current + 1}/{TIEntity.TE_ImageBufferColl.Count}";
    }
}
