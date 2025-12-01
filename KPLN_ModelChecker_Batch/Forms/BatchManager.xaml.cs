using Autodesk.Revit.UI;
using KPLN_ModelChecker_Batch.Forms.Entities;
using NLog;
using System.Media;
using System.Windows;
using System.Windows.Input;

namespace KPLN_ModelChecker_Batch.Forms
{
    public partial class BatchManager : Window
    {
        public BatchManager(Logger currentLogger, UIApplication uiapp)
        {
            InitializeComponent();

            CurrentModel =  new BatchManagerModel(this, currentLogger, uiapp);

            DataContext = CurrentModel;

            PreviewKeyDown += new KeyEventHandler(HandleEsc);
            SystemSounds.Beep.Play();
        }

        public BatchManagerModel CurrentModel { get; }

        /// <summary>
        /// Обработка прокрутки колесика на части окна с файлами
        /// </summary>
        private void FileWrapPanel_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Передаём событие ScrollViewer'у
            if (FileStackScroll != null)
            {
                // Если крутим вниз, увеличиваем вертикальное смещение
                if (e.Delta < 0)
                    FileStackScroll.ScrollToVerticalOffset(FileStackScroll.VerticalOffset + 20); // Примерный шаг прокрутки
                // Если крутим вверх, уменьшаем вертикальное смещение
                else if (e.Delta > 0)
                    FileStackScroll.ScrollToVerticalOffset(FileStackScroll.VerticalOffset - 20);

                // Указываем, что событие обработано
                e.Handled = true;
            }
        }

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
