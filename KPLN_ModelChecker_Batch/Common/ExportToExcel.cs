using KPLN_Library_Forms.UI;
using KPLN_ModelChecker_Batch.Forms.Entities;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_Lib.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace KPLN_ModelChecker_Batch.Common
{
    /// <summary>
    /// Обертка основных данных проверки для каждого проверяемого файла.
    /// </summary>
    internal struct ExcelDataEntity
    {
        internal string CheckRunData { get; set; }

        internal string FileName { get; set; }

        internal string FileSize { get; set; }

        internal int CountedLinks { get; set; }

        internal int OpenedLinks { get; set; }

        internal int CountedWorksets { get; set; }

        internal int OpenedWorksets { get; set; }

        internal List<CheckData> ChecksData { get; set; }
    }

    /// <summary>
    /// Обертка результатов проверки для каждого проверяемого файла.
    /// ВАЖНО: Напрямую ссылку на CheckEntity делать нельзя, т.к. все внутренности перезаписываются по новому открытому файлу. Только через эту оболочку!
    /// </summary>
    internal struct CheckData
    {
        internal string PluginName { get; set; }
        
        /// <summary>
        /// Копия данных по результатам проверки (напрямую из проверки - НЕ ЗАБИРАТЬ)
        /// </summary>
        internal CheckerEntity[] PluginCheckerEntitiesColl { get; set; }

        internal CheckResultStatus CheckRunResult { get; set; }
    }

    /// <summary>
    /// Экспортер в Excel
    /// </summary>
    internal class ExportToExcel
    {
        // HRESULT 0x800A03EC
        private const int _hr = -2146827284;

        /// <summary>
        /// Запуск окна выбора пути для сохранения файла
        /// </summary>
        /// <returns>Путь</returns>
        public static string SetPath()
        {
            // Create an instance of FolderBrowserDialog
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog()
            {
                Description = "Укажите путь для сохранения будущего отчета",
            };

            // Show the dialog and capture the result
            DialogResult result = folderBrowserDialog.ShowDialog();

            if (result == DialogResult.OK) return folderBrowserDialog.SelectedPath;
            else return null;
        }

        /// <summary>
        /// Запуск процесса записи в Excel файл
        /// </summary>
        /// <param name="path">Путь</param>
        /// <param name="excelEntities">Коллекция структуры-оболочки</param>
        public static string Run(string path, ExcelDataEntity[] excelEntities)
        {
            // Имя проверок с кастомным набором полей
            string checkLinksName = new CheckLinks().PluginName;

            Excel.Application excelApp = new Excel.Application
            {
                SheetsInNewWorkbook = excelEntities.Length
            };
            Excel.Workbook workbook = excelApp.Workbooks.Add();

            for (int i = 0; i < excelEntities.Length; i++)
            {
                ExcelDataEntity excelEnt = excelEntities[i];

                Excel.Worksheet worksheet = workbook.Sheets[i + 1];

                string wsName;
                // max 31 символ
                if (excelEnt.FileName.Length > 31)
                {
                    wsName = excelEnt.FileName.Remove(28);
                    wsName = $"{wsName}...";
                }
                else
                    wsName = excelEnt.FileName;
                worksheet.Name = wsName;

                // Настройки общего внешнего вида таблицы
                worksheet.Cells.WrapText = true;
                worksheet.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                worksheet.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                worksheet.Cells.NumberFormat = "@";
                worksheet.Rows.AutoFit();

                // Заполнить заголовки и общие сведения
                worksheet.Columns[1].ColumnWidth = 30;
                worksheet.Cells[1, 1].Value = "ОБЩИЕ СВЕДЕНИЯ:";
                worksheet.Cells[1, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                worksheet.Cells[1, 1].Font.Bold = true;
                worksheet.Cells[2, 1].Value = "Дата/время";
                worksheet.Cells[2, 2].Value = excelEnt.CheckRunData;
                worksheet.Cells[3, 1].Value = "Название файла";
                worksheet.Cells[3, 2].Value = excelEnt.FileName;
                worksheet.Cells[4, 1].Value = "Связей открыто/Всего связей";
                worksheet.Cells[4, 2].Value = $"{excelEnt.OpenedLinks}/{excelEnt.CountedLinks}";
                worksheet.Cells[5, 1].Value = "Рабочих наборов открыто/Всего рабочих наборов";
                worksheet.Cells[5, 2].Value = $"{excelEnt.OpenedWorksets}/{excelEnt.CountedWorksets}";
                worksheet.Cells[6, 1].Value = "Размер файла [МБ]";
                worksheet.Cells[6, 2].Value = excelEnt.FileSize;

                // Записать данные по проверкам
                worksheet.Cells[8, 1].Value = "ВЫЯВЛЕННЫЕ ОШИБКИ:";
                worksheet.Cells[8, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                worksheet.Cells[8, 1].Font.Bold = true;

                int row = 9;
                // Пометка для дополнительной строки ошибки открытых линков
                bool allLinksOpened = excelEnt.OpenedLinks == excelEnt.CountedLinks;
                foreach (CheckData chData in excelEnt.ChecksData)
                {
                    // Заголовки данных по проверкам (!!!если добавишь строку - не забудь поменять кол-во строк для смещения в конце!!!)
                    worksheet.Cells[row, 1].Value = "Имя проверки";
                    worksheet.Cells[row, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                    worksheet.Cells[row, 2].Value = chData.PluginName;
                    worksheet.Cells[row, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);

                    // Кастомный заголовок для проверки связей
                    if (!allLinksOpened && chData.PluginName == checkLinksName)
                    {
                        row++;
                        worksheet.Cells[row, 1].Value = "!!!Внимание!!!";
                        worksheet.Cells[row, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        worksheet.Cells[row, 2].Value = "Не все связи удалось загрузить. Нужно вмешаться человеку";
                        worksheet.Cells[row, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }

                    // Кастомный заголовок для проверок с критической ошибкой запуска
                    if (chData.CheckRunResult == CheckResultStatus.Failed)
                    {
                        row++;
                        worksheet.Cells[row, 1].Value = "!!!Внимание!!!";
                        worksheet.Cells[row, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        worksheet.Cells[row, 2].Value = "Проверку НЕВОЗМОЖНО запустить автоматически. Нужен ручной запуск пользователем, чтобы исправить критические ошибки при запуске";
                        worksheet.Cells[row, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        row += 2;

                        // Выхожу из цикла, т.к. показывать в отчете нечего
                        continue;
                    }

                    // Кастомный заголовок для проверок с ошибкой, при которой НЕТ элементов для проверки
                    if (chData.CheckRunResult == CheckResultStatus.NoItemsToCheck)
                    {
                        row++;
                        worksheet.Cells[row, 1].Value = "!!!Внимание!!!";
                        worksheet.Cells[row, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        worksheet.Cells[row, 2].Value = "Проверку НЕВОЗМОЖНО запустить, т.к. в модели нет подходящих элементов";
                        worksheet.Cells[row, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        row += 2;

                        // Выхожу из цикла, т.к. показывать в отчете нечего
                        continue;
                    }


                    Dictionary<string, string> checkerEntityData = GetStringFromat(chData);

                    int column = 2;
                    foreach (var kvp in checkerEntityData)
                    {
                        worksheet.Columns[column].ColumnWidth = 100;


                        worksheet.Cells[row + 1, 1].Value = "Описание ошибки";
                        worksheet.Cells[row + 1, column].Value = kvp.Key;


                        worksheet.Cells[row + 2, 1].Value = "ID/Имена элементов";
                        worksheet.Cells[row + 2, column].Value = kvp.Value;


                        column++;
                    }

                    // Добавляется кол-во строк
                    row += 4;
                }
            }

            // Сохранить файл
            string currentPath = $"{path}\\Отчет по ошибкам_{ReplaceSpecialCharacters(DateTime.Now.ToString("t"))}.xlsx";

            CustomMessageBox msgToUser = null;
            try
            {
                workbook.SaveAs(currentPath);

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                if (ex.HResult == _hr)
                    msgToUser = new CustomMessageBox("Ошибка!", "Файл занят. Закрой его, либо сохрани файл с другим именем (ищи появившееся окно Excel)");
                else
                    msgToUser = new CustomMessageBox("Ошибка!", "Отправь имя файла и имя проверки проверку разработчику");

            }
            catch (Exception ex)
            {
                msgToUser = new CustomMessageBox("Ошибка!", $"Отправь имя файла и имя проверки проверку разработчику. Текст ошибки: {ex.Message}");
            }
            finally
            {
                // Вывод сообщения пользователю
                msgToUser?.Show();

                // Очистка
                workbook.Close();
                excelApp.Quit();
            }

            return currentPath;
        }

        /// <summary>
        /// Проверка текста на наличие запрещенных символов и замена их на "_"
        /// </summary>
        private static string ReplaceSpecialCharacters(string input)
        {
            char[] specialCharacters = { '<', '>', '?', '[', ']', ':', '|' };
            foreach (char c in specialCharacters)
            {
                input = input.Replace(c, '_');
            }
            return input;
        }

        /// <summary>
        /// Форматирование данных из ошибок в нужный текстовый формат
        /// </summary>
        /// <returns>1 - Описание ошибок; 2 - ID/Имя-элементов </returns>
        private static Dictionary<string, string> GetStringFromat(CheckData chData)
        {
            CheckerEntity[] docErrors = chData.PluginCheckerEntitiesColl;

            // Если ошибок не выявили
            if (chData.CheckRunResult == CheckResultStatus.Succeeded && docErrors.Length == 0)
                return new Dictionary<string, string>()
                {
                    { "Ошибок не выявлено!", "-" }
                };


            // Имя проверок-исключений, для которых вместо ID - нужен кастом
            string checkFamName = new CheckFamilies().PluginName;


            // Генерирую словарь
            Dictionary<string, string> dict = new Dictionary<string, string>();
            foreach (CheckerEntity entity in docErrors)
            {
                string errorStatus = string.Empty;
                switch (entity.Status)
                {
                    case ErrorStatus.Error:
                        errorStatus = "Критическая ошибка";
                        break;
                    case ErrorStatus.Warning:
                        errorStatus = "Предупреждение";
                        break;
                    case ErrorStatus.LittleWarning:
                        goto case ErrorStatus.AllmostOk;
                    case ErrorStatus.AllmostOk:
                        errorStatus = "Для справки";
                        break;
                    case ErrorStatus.Approve:
                        errorStatus = "Допустимое";
                        break;
                }

                string approveComment = string.Empty;
                if (entity.Status == ErrorStatus.Approve)
                    approveComment = $"[{entity.ApproveComment}]";


                string[] headerDataParts = new string[5]
                    {
                        $"[{errorStatus}]",
                        approveComment,
                        entity.Header,
                        entity.Description,
                        entity.Info,
                    }
                    // Чтобы не было пустых стрелок
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToArray();

                string headerData = string.Join("→", headerDataParts);
                string IdData = string.Empty;
                if (chData.PluginName == checkFamName)
                    IdData = string.Join(", ", entity.ElementName);
                else
                    IdData = string.Join(", ", entity.ElementIdCollection.ToList());


                if (dict.TryGetValue(headerData, out string dictIdData))
                    dict[headerData] = string.Concat(dictIdData, ", ", IdData);
                else
                    dict[headerData] = IdData;
            }

            return dict;
        }
    }
}
