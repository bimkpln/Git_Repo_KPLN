using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;
using System.Linq;
using System;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_Tools.Forms
{
    /// <summary>
    /// Логика взаимодействия для AR_TEPDesign_categorySelect.xaml
    /// </summary>
    public partial class AR_TEPDesign_categorySelect : Window
    {
        public int Result { get; private set; } = 0;
        private UIDocument uidoc;
        private ViewSheet sheet;

        public AR_TEPDesign_categorySelect(UIDocument _uidoc, ViewSheet _sheet)
        {
            InitializeComponent();
            uidoc = _uidoc;
            sheet = _sheet;
        }

        private void CategoryButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && int.TryParse(button.Tag?.ToString(), out int value))
            {
                Result = value;
                Close();
            }
        }

        private void UpdateTablleButton_Click(object sender, RoutedEventArgs e)
        {
            Result = 5;

            Document doc = uidoc.Document;
            string prefix = $"ТЭП_{sheet.SheetNumber}";
            string suffix = "_update";


            var allSchedules = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>();
            var schedulesToAdd = allSchedules
                .Where(vs =>
                    vs.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                    && !vs.Name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
                )
                .ToList();

            if (schedulesToAdd.Count == 0)
            {
                this.Close();
                TaskDialog.Show(
                    "Информация",
                    $"Не найдено ни одной спецификации, которую необходимо обновить.");               
                return;
            }

            int countAdded = 0;
            using (Transaction tUST = new Transaction(doc, "KPLN. ТЭП. Обновление стиля таблицы"))
            {
                tUST.Start();

                try
                {
                    foreach (ViewSchedule vs in schedulesToAdd)
                    {
                        vs.Name = vs.Name + suffix;
                        countAdded++;
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Ошибка", $"Возникла ошибка при обновлении стиля спецификации:\n{ex}.");
                    tUST.RollBack();
                    return;
                }

                tUST.Commit();
            }

            var schedulesToRemove = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs =>
                    vs.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                    && vs.Name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
                )
                .ToList();

            int countRemoved = 0;
            using (Transaction tUV = new Transaction(doc, "KPLN. ТЭП. Обновление данных на странице"))
            {
                tUV.Start();

                try
                {
                    foreach (ViewSchedule vs in schedulesToRemove)
                    {
                        string newName = vs.Name.Substring(0, vs.Name.Length - suffix.Length);
                        vs.Name = newName;
                        countRemoved++;
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Ошибка", $"Возникла ошибка при обновлении стиля спецификации:\n{ex}.");
                    tUV.RollBack();
                    return;
                }

                tUV.Commit();
            }

            Close();
            TaskDialog.Show("Успех",$"Операция завершена.");         
        }
    }
}
