using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace KPLN_ModelChecker_Batch.Common
{
    internal struct ExcelDataEntity
    {
        internal string CheckRunData;

        internal string FileName;

        internal string FileSize;

        internal string OpenedLinks;

        internal string OpenedWorksets;

        internal List<CheckData> ChecksData;
    }

    internal struct CheckData
    {
        internal string PluginName;

        internal CheckerEntity[] CheckEntities;
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
        public static void Run(string path, ExcelDataEntity[] excelEntities)
        {
            Excel.Application excelApp = new Excel.Application
            {
                SheetsInNewWorkbook = excelEntities.Length
            };
            Excel.Workbook workbook = excelApp.Workbooks.Add();

            for (int i = 0; i < excelEntities.Length; i++)
            {
                ExcelDataEntity excelEnt = excelEntities[i];

                Excel.Worksheet worksheet = workbook.Sheets[i+1];
                worksheet.Name = excelEnt.FileName;

                // Настройки внешнего вида таблицы
                worksheet.Columns[1].ColumnWidth = 30;
                worksheet.Cells.WrapText = true;
                worksheet.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                worksheet.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                worksheet.Cells.NumberFormat = "@";
                worksheet.Rows.AutoFit();

                // Заполнить заголовки и общие сведения
                worksheet.Cells[1, 1].Value = "ОБЩИЕ СВЕДЕНИЯ:";
                worksheet.Cells[2, 1].Value = "Дата/время";
                worksheet.Cells[2, 2].Value = excelEnt.CheckRunData;
                worksheet.Cells[3, 1].Value = "Название файла";
                worksheet.Cells[3, 2].Value = excelEnt.FileName;
                worksheet.Cells[4, 1].Value = "Связей открыто/Всего связей";
                worksheet.Cells[4, 2].Value = excelEnt.OpenedLinks;
                worksheet.Cells[5, 1].Value = "Рабочих наборов открыто/Всего рабочих наборов";
                worksheet.Cells[5, 2].Value = excelEnt.OpenedWorksets;
                worksheet.Cells[6, 1].Value = "Размер файла [МБ]";
                worksheet.Cells[6, 2].Value = excelEnt.FileSize;

                // Записать данные по проверкам
                worksheet.Cells[8, 1].Value = "ВЫЯВЛЕННЫЕ ОШИБКИ:";
                int row = 9;
                foreach (CheckData chData in excelEnt.ChecksData)
                {
                    // Заголовки данных по проверкам (!!!если добавишь строку - не забудь поменять кол-во строк для смещения в конце!!!)
                    worksheet.Cells[row, 1].Value = "Имя проверки";
                    worksheet.Cells[row, 2].Value = chData.PluginName;
                    worksheet.Cells[row + 1, 1].Value = "Имя элемента/-ов";
                    worksheet.Cells[row + 2, 1].Value = "ID элемента/-ов";
                    worksheet.Cells[row + 3, 1].Value = "Заголовок ошибки";
                    worksheet.Cells[row + 4, 1].Value = "Описание ошибки";
                    worksheet.Cells[row + 5, 1].Value = "Дополнительная информация";


                    // Коллекция полей класса, которые нужно извлечь
                    List<string> selectedFields = new List<string>(6)
                    {
                        nameof(CheckerEntity.ElementName),
                        nameof(CheckerEntity.ElementIdCollection),
                        nameof(CheckerEntity.Header),
                        nameof(CheckerEntity.Description),
                        nameof(CheckerEntity.Info),
                    };

                    int subColumn = 2;
                    foreach (CheckerEntity checkerEnt in chData.CheckEntities)
                    {
                        if (checkerEnt.Status == ErrorStatus.Approve) continue;
                        
                        int subRow = row;
                        worksheet.Columns[subColumn].ColumnWidth = 100;

                        string valueFromColl = string.Empty;
                        foreach (string field in selectedFields)
                        {
                            subRow++;
                            object value = checkerEnt.GetType().GetProperty(field).GetValue(checkerEnt);

                            // Запись коллекции ElementIdCollection
                            if (field == nameof(WPFEntity.ElementIdCollection) && value is IEnumerable<ElementId> list)
                            {
                                valueFromColl = string.Join(", ", list);
                                worksheet.Cells[subRow, subColumn].Value = valueFromColl;
                            }
                            // Запись остальных элементов, если они не равны null
                            else if (value != null)
                                worksheet.Cells[subRow, subColumn].Value = value;
                        }
                        subColumn++;
                    }

                    // Добавляется кол-во строк
                    row += 7;
                }
            }

            // Сохранить файл
            string currentPath = $"{path}\\Отчет по ошибкам_{ReplaceSpecialCharacters(DateTime.Now.ToString("t"))}.xlsx";

            CustomMessageBox customMessageBox = null;
            try
            {
                workbook.SaveAs(currentPath);
                customMessageBox = new CustomMessageBox("Сохранено успешно!", $"Путь:\n{currentPath}");

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                if (ex.HResult == _hr)
                    customMessageBox = new CustomMessageBox("Ошибка!", "Файл занят. Закрой его, либо сохрани файл с другим именем (ищи появившееся окно Excel)");
                else
                    customMessageBox = new CustomMessageBox("Ошибка!", "Отправь имя файла и имя проверки проверку разработчику");

            }
            catch (Exception ex)
            {
                customMessageBox = new CustomMessageBox("Ошибка!", $"Отправь имя файла и имя проверки проверку разработчику. Текст ошибки: {ex.Message}");
            }
            finally
            {
                // Вывод сообщения пользователю
                customMessageBox.ShowDialog();

                // Очистка
                workbook.Close();
                excelApp.Quit();
            }
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
    }
}
