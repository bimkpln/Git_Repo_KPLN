using System.Windows.Forms;

namespace KPLN_Loader.Services
{
    internal static class LoaderMessageService
    {
        internal const string ErrorTitle = "Ошибка KPLN";

        internal static class UserMessages
        {
            internal const string DbConnectionFailedPrefix = "Подключение к БД не состоялось. Проверьте подключение к интернету.";
            internal const string ContactSupport = "Если подключение к интернету в порядке - отправьте письмо на почту \"bim@kpln.ru\" со скриншотом проблемы.";
            internal const string InitializationFailed = "Инициализация не удалась. Проверьте подключение к интернету.";
            internal const string RestrictedAccess =
                "Вам закрыт доступ к плагинам KPLN. Удалите плагин с компьютера, или отправьте запрос на почту \"bim@kpln.ru\".\n\n" +
                "Для удаления перейдите пройдите по пути \"Панель управления\" -> \"Программы\" -> \"Удаление программы\".\n" +
                "Далее поиском найдите \"KPLN_ExtraNet\" и нажмите \"Удалить\".";
            internal const string GlobalInitializationError = "Инициализация не удалась. Отправь в BIM-отдел KPLN: {0}";
        }

        internal static class ModuleMessages
        {
            internal const string MissingModule = "Модуль/библиотека [{0}] не найден/а!";
            internal const string CopyFailed = "Модуль/библиотека {0} - не скопировался по подготовленному пути";
            internal const string MissingDll = "Модуль/библиотека {0} - не активирован. Отсутствует dll для активации";
            internal const string LocalLoadError = "Локальная ошибка загрузки модуля {0}";
            internal const string LibraryLoaded = "Модуль-библиотека [{0}] версии {1} успешно {2}!";
            internal const string ModuleLoaded = "Модуль [{0}] версии {1} успешно активирован!";
            internal const string ModuleNotActivated = "Модуль [{0}] не активирован.";
        }

        internal static void ShowWarning(string text)
        {
            MessageBox.Show(text, ErrorTitle, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
        }
    }
}
