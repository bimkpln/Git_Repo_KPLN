﻿using KPLN_Library_Forms.UI;
using KPLN_ModelChecker_Lib;
using System;
using System.Collections.Generic;
using System.Linq;
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

        internal CheckerEntity[] CheckerEntities;
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
            Excel.Application excelApp = new Excel.Application
            {
                SheetsInNewWorkbook = excelEntities.Length
            };
            Excel.Workbook workbook = excelApp.Workbooks.Add();

            for (int i = 0; i < excelEntities.Length; i++)
            {
                ExcelDataEntity excelEnt = excelEntities[i];

                Excel.Worksheet worksheet = workbook.Sheets[i + 1];
                worksheet.Name = excelEnt.FileName;

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
                worksheet.Cells[1, 1].Font.Bold= true;
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
                worksheet.Cells[8, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                worksheet.Cells[8, 1].Font.Bold = true;
                int row = 9;
                foreach (CheckData chData in excelEnt.ChecksData)
                {
                    // Заголовки данных по проверкам (!!!если добавишь строку - не забудь поменять кол-во строк для смещения в конце!!!)
                    worksheet.Cells[row, 1].Value = "Имя проверки";
                    worksheet.Cells[row, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                    worksheet.Cells[row, 2].Value = chData.PluginName;
                    worksheet.Cells[row, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);

                    Dictionary<string, string> checkerEntityData = GetStringFromat(chData.CheckerEntities);

                    int column = 2;
                    foreach(var kvp in checkerEntityData)
                    {
                        worksheet.Columns[column].ColumnWidth = 100;

                        worksheet.Cells[row + 1, 1].Value = "Описание ошибки";
                        worksheet.Cells[row + 1, column].Value = kvp.Key;


                        worksheet.Cells[row + 2, 1].Value = "ID элемента/-ов";
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
        /// <param name="docErrors"></param>
        /// <returns>1 - Описание ошибок; 2 - ID-элементов </returns>
        private static Dictionary<string, string> GetStringFromat(CheckerEntity[] docErrors)
        {
            if (docErrors == null || docErrors.Length == 0)
            {
                return
                    new Dictionary<string, string>()
                    {
                        { "Ошибок не выявлено!", "-" },
                    };
            }

            Dictionary<string, string> dict = new Dictionary<string, string>();
            foreach (CheckerEntity entity in docErrors)
            {
                string[] headerDataParts = new string[3]
                    {
                        entity.Header,
                        entity.Description,
                        entity.Info,
                    }
                    // Чтобы не было пустых стрелок
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToArray(); 

                string headerData = string.Join("→", headerDataParts);
                string IdData = string.Join(", ", entity.ElementIdCollection.ToList());

                if (dict.TryGetValue(headerData, out string dictIdData))
                    dict[headerData] = string.Concat(dictIdData, ", ", IdData);
                else
                    dict[headerData] = IdData;
            }

            return dict;
        }
    }
}
