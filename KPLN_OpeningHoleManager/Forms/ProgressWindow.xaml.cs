using KPLN_OpeningHoleManager.Forms.MVVMCore_MainMenu;

namespace KPLN_OpeningHoleManager.Forms
{
    /// <summary>
    /// Создаёт библиотечное окно напрямую: WPF загружает BAML из сборки типа окна.
    /// </summary>
    public sealed class ProgressWindow
    {
        private readonly KPLN_Library_Forms.UI.ProgressWindow _window;

        public ProgressWindow(ProgressInfoViewModel progressInfo)
        {
            _window = new KPLN_Library_Forms.UI.ProgressWindow(progressInfo.LibraryProgress)
            {
                Title = "KPLN: Менеджер отверстий"
            };
        }

        public void Show() => _window.Show();

        // Ранний выход из операции должен закрывать окно без запроса отмены.
        public void Close() => _window.CloseSafely();
    }
}
