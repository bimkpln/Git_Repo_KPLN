using KPLN_Library_Forms.UI;
using KPLN_TaskManager.Common;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace KPLN_TaskManager.Services
{
    /// <summary>
    /// Экспортер в Excel
    /// </summary>
    internal class ExportToExcelService
    {
        /// <summary>
        /// HRESULT 0x800A03EC
        /// </summary>
        private const int _hr = -2146827284;

        /// <summary>
        /// Запуск окна выбора пути для сохранения файла
        /// </summary>
        /// <returns>Путь</returns>
        public static string SetPath()
        {
            // Create an instance of FolderBrowserDialog
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            // Show the dialog and capture the result
            DialogResult result = folderBrowserDialog.ShowDialog();

            if (result == DialogResult.OK)
                return folderBrowserDialog.SelectedPath;
            else
                return null;
        }

        /// <summary>
        /// Запуск процесса записи в Excel файл
        /// </summary>
        /// <param name="path">Путь</param>
        /// <param name="prjName">Имя проекта</param>
        /// <param name="entities">Коллекция TaskItemEntity</param>
        public static void Run(string path, string prjName, IEnumerable<TaskItemEntity> entities)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets[1];

            // Коллекция полей класса, которые нужно извлечь
            List<string> selectedFields = new List<string>(6)
            {
                nameof(TaskItemEntity.TaskTitle),
                nameof(TaskItemEntity.TaskStatus),
                nameof(TaskItemEntity.TaskBody),
                nameof(TaskItemEntity.CreatedTaskUserFullName),
                nameof(TaskItemEntity.LastChangeData),
            };

            // Заполнить заголовки столбцы (должны мапиться со списком полей класса) 
            worksheet.Cells[1, 1].Value = "ИМЯ ЗАДАЧИ";
            worksheet.Cells[1, 2].Value = "СТАТУС ЗАДАЧИ";
            worksheet.Cells[1, 3].Value = "ОПИСАНИЕ ЗАДАЧИ";
            worksheet.Cells[1, 4].Value = "ДАННЫЕ О СОЗДАТЕЛЕ";
            worksheet.Cells[1, 5].Value = "ДАННЫЕ ОБ ИЗМЕНЕНИЯХ В ЗАДАЧЕ";

            // Записать данные
            int row = 2;
            foreach (TaskItemEntity entity in entities)
            {
                int column = 1;
                string valueFromColl = string.Empty;
                foreach (string field in selectedFields)
                {
                    object value = entity.GetType().GetProperty(field).GetValue(entity);
                    if (value != null)
                    {
                        if (field.Equals(nameof(TaskItemEntity.TaskStatus)))
                        {
                            TaskStatusEnum taskStatusEnum = (TaskStatusEnum)value;
                            switch (taskStatusEnum)
                            {
                                case TaskStatusEnum.Open:
                                    value = "Открыто";
                                    break;
                                case TaskStatusEnum.Close:
                                    value = "Закрыто";
                                    break;
                            }
                        }

                        worksheet.Cells[row, column++].Value = value;
                        // Рисунок можно вставить в ячейку не для всех версий эксель. Заглушено
                        //if (field.Equals(nameof(TaskItemEntity.ImageBuffer)))
                        //{
                        //    byte[] imageBuffer = (byte[])value;
                        //    InsertImageToExcel(worksheet, row, column++, imageBuffer);
                        //}
                        //else
                        //    worksheet.Cells[row, column++].Value = value;
                    }
                }
                row++;
            }

            // Сохранить файл
            string currentPath = $"{path}\\{ReplaceSpecialCharacters(prjName)}_Отчет по задачам.xlsx";

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
                customMessageBox.Topmost = true;
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

        private static void InsertImageToExcel(Excel.Worksheet worksheet, int row, int column, byte[] imageBytes)
        {
            if (imageBytes == null || imageBytes.Length == 0) return;

            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                string tempPath = Path.GetTempFileName();
                img.Save(tempPath, System.Drawing.Imaging.ImageFormat.Png);

                worksheet.Shapes.AddPicture(tempPath,
                    MsoTriState.msoFalse,
                    MsoTriState.msoCTrue,
                    worksheet.Cells[row, column].Left,
                    worksheet.Cells[row, column].Top,
                    img.Width / 2,
                    img.Height / 2);

                File.Delete(tempPath);
            }
        }
    }
}
