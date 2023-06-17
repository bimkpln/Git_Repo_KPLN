using Autodesk.Revit.DB;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Экспортер в Excel
    /// </summary>
    internal static class WPFEntity_ExportToExcel
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
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            // Show the dialog and capture the result
            DialogResult result = folderBrowserDialog.ShowDialog();

            if (result == DialogResult.OK) return folderBrowserDialog.SelectedPath;
            else return null;
        }

        /// <summary>
        /// Запуск процесса записи в Excel файл
        /// </summary>
        /// <param name="path">Путь</param>
        /// <param name="checkName">Имя проверки</param>
        /// <param name="entities">Коллекция WPFEntity</param>
        public static void Run(string path, string checkName, IEnumerable<WPFEntity> entities)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets[1];

            // Коллекция полей класса, которые нужно извлечь
            List<string> selectedFields = new List<string>(6)
            {
                nameof(WPFEntity.ElementName),
                nameof(WPFEntity.ElementId),
                nameof(WPFEntity.ElementIdCollection),
                nameof(WPFEntity.ErrorHeader),
                nameof(WPFEntity.Description),
                nameof(WPFEntity.Info),
            };

            // Заполнить заголовки столбцы (должны мапиться со списком полей класса) 
            worksheet.Cells[1, 1].Value = "ИМЯ ЭЛЕМЕНТА";
            worksheet.Cells[1, 2].Value = "ID ЭЛЕМЕНТА/-ОВ";
            worksheet.Cells[1, 3].Value = "ЗАГОЛОВОК ОШИБКИ";
            worksheet.Cells[1, 4].Value = "ОПИСАНИЕ ОШИБКИ";
            worksheet.Cells[1, 5].Value = "ДОПОЛНИТЕЛЬНАЯ ИНФОРМАЦИЯ ПО ОШИБКЕ";

            // Записать данные
            int row = 2;
            foreach (WPFEntity entity in entities)
            {
                if (entity.CurrentStatus == Collections.Status.Approve) continue;

                int column = 1;
                string valueFromColl = string.Empty;
                foreach (string field in selectedFields)
                {
                    object value = entity.GetType().GetProperty(field).GetValue(entity);
                    // Запись ElementId
                    if (field == nameof(WPFEntity.ElementId) && value == null) continue;

                    // Запись коллекции ElementIdCollection
                    if (field == nameof(WPFEntity.ElementIdCollection) && value is IEnumerable<ElementId> list)
                    {
                        valueFromColl = string.Join(", ", list);
                        worksheet.Cells[row, column++].Value = valueFromColl;
                    }
                    // Запись остальных элементов, если они не равны null
                    else if (value != null) worksheet.Cells[row, column++].Value = value;
                }
                row++;
            }

            // Сохранить файл
            string currentPath = $"{path}\\{checkName}_Отчет по ошибкам.xlsx";
            try
            {
                workbook.SaveAs(currentPath);
                MessageBox.Show($"Сохранено успешно!\nПуть: {currentPath}");
            }
            catch (System.Runtime.InteropServices.COMException ex) 
            {
                if (ex.HResult == _hr) MessageBox.Show("Файл занят. Закрой его, либо сохрани файл с другим именем (ищи появившееся окно Excel)");
                else MessageBox.Show("Ой. Отправь имя файла, проверку разработчику");
            }
            finally
            {
                // Очистка
                workbook.Close();
                excelApp.Quit();
            }
        }
    }
}
