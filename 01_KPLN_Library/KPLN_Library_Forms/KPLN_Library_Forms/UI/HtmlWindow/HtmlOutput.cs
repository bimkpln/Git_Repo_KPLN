using Autodesk.Revit.UI;
using HtmlAgilityPack;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace KPLN_Library_Forms.UI.HtmlWindow
{
    public static class HtmlOutput
    {
        /// <summary>
        /// Маска для имени файлов (приставка)
        /// </summary>
        private static readonly string _fileNameMask = "KPLN_htmlPrint";
        /// <summary>
        /// Путь к папке для хранения файлов
        /// </summary>
        private static readonly string _fileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        /// <summary>
        /// Путь для сохранения файлов
        /// </summary>
        private static readonly string _outputPath = Path.Combine(_fileDirectory,  $"{_fileNameMask}_{DateTime.Now:dd/MM/yyyy_HH/mm/ss}.html");
        private static readonly string HTML_Output_DocType = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>";
        /// <summary>
        /// Флаг для индикации запуска чистки предыдущих версий документов в текущей сессии
        /// </summary>
        private static bool _isCleared = false;
        private static string _html_Output_Head ;
        private static HtmlDocument _htmlDocument = new HtmlDocument();

        public static void Main()
        {
            CreateStream();
        }

        /// <summary>
        /// Окно вывода информации пользователю
        /// </summary>
        public static OutputWindow FormOutput { get; set; }

        /// <summary>
        /// Настрока html-файла
        /// </summary>
        private static string HTML_Output_Head 
        {
            get
            {
                if (_html_Output_Head == null)
                {
                    string resCssPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString();
                    _html_Output_Head = string.Format(@"<head><meta http-equiv='X-UA-Compatible' content='IE=9'><meta http-equiv='content-type' content='text/html; charset=utf-8'><meta name='appversion' content='0.2.0.0'><link href='file:///{0}\UI\HtmlWindow\Styles\outputstyles.css' rel='stylesheet'></head>", resCssPath).ToString();
                }

                return _html_Output_Head;
            }
        }

        /// <summary>
        /// Принтер
        /// </summary>
        /// <param name="value">Строка</param>
        /// <param name="cssclass">Стиль css</param>
        private static void OutputPrint(string value, string cssclass)
        {
            // Предварительная чистка старых файлов
            if (!_isCleared)
            {
                Task clearingTask = Task.Run(() => ClearOldFiles());
            }
            
            // Печать
            string insertablevalue = string.Join("<b>«", value.Split('«'));
            insertablevalue = string.Join("»</b>", insertablevalue.Split('»'));
            insertablevalue = string.Join("<b>[", insertablevalue.Split('['));
            insertablevalue = string.Join("]</b>", insertablevalue.Split(']'));

            if (FormOutput == null)
            {
                FormOutput = new OutputWindow();
                CreateStream();
                FormOutput.Show();
                FormOutput.webBrowser.Navigate(new Uri(_outputPath));
            }
            else
            {
                FormOutput.Show();
            }

            HtmlNode body = _htmlDocument.DocumentNode.SelectSingleNode("/html/body");
            if (body == null)
            {
                CreateStream();
                body = _htmlDocument.DocumentNode.SelectSingleNode("/html/body");
            }

            HtmlNode node = HtmlNode.CreateNode(string.Format("<p class='{0}'>{1}</p>", cssclass, insertablevalue));
            body.InsertBefore(node, body.FirstChild);
            using (StreamWriter streamWriter = new StreamWriter(_outputPath))
            {
                _htmlDocument.Save(streamWriter);
                streamWriter.Close();
            }

            FormOutput.webBrowser.Refresh();
            System.Windows.Forms.Application.DoEvents();
            FormOutput.webBrowser.Refresh();
            FormOutput.BringToFront();
        }

        private static void CreateStream()
        {
            _htmlDocument = new HtmlDocument();
            _htmlDocument.LoadHtml(string.Format(@"<html>{0}{1}<body></body></html>", HTML_Output_DocType, HTML_Output_Head));

            using (StreamWriter streamWriter = new StreamWriter(_outputPath))
            {
                _htmlDocument.Save(streamWriter);
                streamWriter.Close();
            }
        }

        /// <summary>
        /// Метод очистки от файлов пердыдущего запуска
        /// </summary>
        private static void ClearOldFiles()
        {
            // Перевожу в сигнальное состояние, чтобы больше чистку не запускать
            _isCleared = true;

            // Провожу чистку
            string[] htmlOutputFilesFullPath = Directory.GetFiles(_fileDirectory).Where(s => s.Contains(_fileNameMask)).ToArray();
            foreach (string fullPath in htmlOutputFilesFullPath)
            {
                FileInfo file = new FileInfo(fullPath);
                if (file.CreationTime.Date < DateTime.Now.AddDays(-3))
                {
                    try
                    {
                        file.Delete();
                    }
                    // Ошибка будет только если файл занят
                    catch (UnauthorizedAccessException) { }
                    catch (Exception ex)
                    {
                        TaskDialog td = new TaskDialog("KPLN")
                        {
                            MainInstruction = "Скинь скрин ошибки в BIM-отдел",
                            MainContent = $"При очистке старых файлов - произошла ошика: {ex.Message}"
                        };
                        td.Show();
                    }
                }
            }
        }

        /// <summary>
        /// Вывод чистой ошибки
        /// </summary>
        /// <param name="e">Экземпля класса Exception</param>
        public static void PrintError(Exception e)
        {
            OutputPrint(e.StackTrace, "code");
            OutputPrint(e.Message, "logerrorheader");
        }

        /// <summary>
        /// Вывод ошибки с дополнительным описанием
        /// </summary>
        /// <param name="e">Экземпля класса Exception</param>
        /// <param name="message">Дополнительное описание</param>
        public static void PrintError(Exception e, string message)
        {
            OutputPrint(e.StackTrace, "code");
            OutputPrint(string.Format("{0}<br>   {1}", message, e.Message), "logerrorheader");
        }

        /// <summary>
        /// Вывод сообщения с оформлением по типу сообщения
        /// </summary>
        /// <param name="value">Строка сообщения</param>
        /// <param name="type">Тип ошибки</param>
        public static void Print(string value, MessageType type)
        {
            switch (type)
            {
                case MessageType.Error:
                    OutputPrint(value, "logerror");
                    break;
                case MessageType.Header:
                    OutputPrint(value, "logheader");
                    break;
                case MessageType.Success:
                    OutputPrint(value, "logsuccess");
                    break;
                case MessageType.Warning:
                    OutputPrint(value, "logwarning");
                    break;
                case MessageType.Critical:
                    OutputPrint(value, "logcritical");
                    break;
                case MessageType.Code:
                    OutputPrint(value, "code");
                    break;
                case MessageType.System_OK:
                    OutputPrint(value, "systemok");
                    break;
                case MessageType.System_Regular:
                    OutputPrint(value, "systemregular");
                    break;
                default:
                    OutputPrint(value, "logdefault");
                    break;
            }
        }
    }
}
