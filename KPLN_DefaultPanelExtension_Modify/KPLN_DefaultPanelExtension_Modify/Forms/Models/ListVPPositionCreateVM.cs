using Autodesk.Revit.DB;
using KPLN_DefaultPanelExtension_Modify.ExecutableCommands;
using System;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace KPLN_DefaultPanelExtension_Modify.Forms.Models
{
    public sealed class ListVPPositionCreateVM
    {
        private readonly Window _owner;
        private readonly Element _selVE;

        public ListVPPositionCreateVM(Window owner, Element selVE)
        {
            _owner = owner;
            _selVE = selVE;

            SelectAlignCmd = new RelayCommand<object>(SelectAlign);
            SaveCmd = new RelayCommand<object>(Save);
            HelpCmd = new RelayCommand<object>(_ => Help());
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);

            ListVPPositionItem = new ListVPPositionCreateM(_selVE);
        }

        public ICommand SelectAlignCmd { get; }

        public ICommand SaveCmd { get; }
        
        public ICommand HelpCmd { get; }

        public ICommand CloseWindowCmd { get; }

        public ListVPPositionCreateM ListVPPositionItem { get; set; }

        private void SelectAlign(object selParamObj)
        {
            if (Enum.TryParse(selParamObj.ToString(), out AlignMode mode))
                ListVPPositionItem.SelectedAlign = mode;
        }

        private void Save(object windObj)
        {
            if (string.IsNullOrWhiteSpace(ListVPPositionItem.ConfigName))
            {
                MessageBox.Show(
                    _owner,
                    $"Не указано имя конфигурации", "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return;
            }

            if (ListVPPositionItem.ConfigName.Length < 3D)
            {
                MessageBox.Show(
                    _owner,
                    $"Имя конфигурации слишком короткое", "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return;
            }

            // Уточняю имя системными приставками
            ListVPPositionItem.ConfigName = $"{SetSystemPartConfigName()}{ListVPPositionItem.ConfigName}";


            if (windObj is Window window)
            {
                window.DialogResult = true;
                window.Close();
            }
        }

        private void Help() =>
            Process.Start(new ProcessStartInfo(ExcCmdListVPPositionStart.HelpUrl) { UseShellExecute = true });

        private void CloseWindow(object windObj)
        {
            ListVPPositionItem = null;

            if (windObj is Window window)
            {
                window.DialogResult = false;
                window.Close();
            }
        }

        private string SetSystemPartConfigName()
        {
            // Размер основной надписи
            string titleNamePart = string.Empty;
            Document doc = _selVE.Document;
            if (doc.ActiveView is ViewSheet vSheet)
            {
                var tBlocks = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    //.Where(el => el.Category != null && el.Category.BuiltInCategory == BuiltInCategory.OST_TitleBlocks)
                    .Cast<FamilyInstance>();

                foreach (var tBlock in tBlocks)
                {
                    Parameter sizeParam = tBlock.LookupParameter("А");
                    Parameter xParam = tBlock.LookupParameter("х");
                    if (sizeParam != null && xParam != null)
                    {
                        titleNamePart = $"[А{sizeParam.AsValueString()}х{xParam.AsValueString()}]_";
                        break;
                    }
                }
            }

            // Выравнивание
            string alignNamePart = string.Empty;
            switch (ListVPPositionItem.SelectedAlign)
            {
                case (AlignMode.LeftTop):
                    alignNamePart = "[ЛВ]: ";
                    break;
                case (AlignMode.RightTop):
                    alignNamePart = "[ПВ]: ";
                    break;
                case (AlignMode.Center):
                    alignNamePart = "[Ц]: ";
                    break;
                case (AlignMode.LeftBottom):
                    alignNamePart = "[ЛН]: ";
                    break;
                case (AlignMode.RightBottom):
                    alignNamePart = "[ПН]: ";
                    break;
                case (AlignMode.OrignToOrigin):
                    alignNamePart = "[СВН]: ";
                    break;
            }


            return $"{titleNamePart}{alignNamePart}";
        }
    }
}
