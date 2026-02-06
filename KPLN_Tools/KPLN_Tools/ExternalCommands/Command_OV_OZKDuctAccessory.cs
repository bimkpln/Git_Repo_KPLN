using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common.OVVK_System;
using KPLN_Tools.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class Command_OV_OZKDuctAccessory : IExternalCommand
    {
        internal const string PluginName = "ОВ: Клапаны ОЗК";

        private static readonly Guid _prefParamGuid = new Guid("3a473dae-91b8-46e7-9db6-8f3371f4a5d4");
        private static readonly Guid _sufParamGuid = new Guid("7cdd0d4b-7252-45a8-bc26-d0e63d871fa9");

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteByUIApp(commandData.Application);
        }

        public Result ExecuteByUIApp(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            if (uidoc == null)
                return Result.Cancelled;


            // Собираю данные из модели
            Document doc = uidoc.Document;
            FamilySymbol[] famSybols = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctAccessory)
                .WhereElementIsElementType()
                .Where(el => string.Compare(el.LookupParameter("Ключевая пометка").AsString(), "KPLN: Универсальный клапан ОЗК", StringComparison.OrdinalIgnoreCase) == 0)
                .Cast<FamilySymbol>()
                .ToArray();


            // Отлов ошибки
            if (famSybols.Length == 0)
            {
                MessageBox.Show(
                    $"В проекте отсутсвуют спец. семейства универсальных клапанов ОЗК (ищу по параметру \"Ключевая пометка\")",
                    "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Asterisk);

                return Result.Cancelled;
            }


            // Подготовка сущностей
            List<OZKDuctAccessoryEntity> ozkEntities = new List<OZKDuctAccessoryEntity>();
            foreach (FamilySymbol fs in famSybols)
            {
                IEnumerable<FamilyInstance> famInsts = fs
                    .GetDependentElements(null)
                    .Where(id => doc.GetElement(id) is FamilyInstance famInst && famInst.SuperComponent == null)
                    .Select(id => (FamilyInstance)doc.GetElement(id));
                
                if (!famInsts.Any())
                    continue;


                // Приставка и суффиксы это параметры экземпляра, НО за их целостность отвечает структура сайзлукапа
                // Эти значения внутри одного типа одинаковые, т.к. они в экзмепляре по вынужденной причине - вылет лопаток зависит от сечения
                // НО при этом в маркировках эта инфа не меняется.
                // Исходя из этого - создаю сущности с одинаковыми значениями параметров суффикса и приставки
                HashSet<string> prefFIData = famInsts.Select(fi => fi.get_Parameter(_prefParamGuid).AsString()).ToHashSet<string>();
                HashSet<string> suffFIData = famInsts.Select(fi => fi.get_Parameter(_sufParamGuid).AsString()).ToHashSet<string>();
                foreach(string pref in prefFIData)
                {
                    IEnumerable<FamilyInstance> prefFI = famInsts.Where(fi => fi.get_Parameter(_prefParamGuid).AsString().Equals(pref));
                    foreach (string suff in suffFIData)
                    {
                        IEnumerable<FamilyInstance> prefAndSuffFI = prefFI.Where(fi => fi.get_Parameter(_sufParamGuid).AsString().Equals(suff));
                        if (!prefAndSuffFI.Any())
                        {
                            MessageBox.Show(
                                $"Отправь разработчику - не удалось объединить данные по приставке и по суффиксу",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Asterisk);

                            return Result.Cancelled;
                        }

                        ozkEntities.Add(new OZKDuctAccessoryEntity(fs, prefAndSuffFI, pref, suff));
                    }
                }
            }


            // Создание формы
            OV_OZKDuctAccessoryForm mainForm = new OV_OZKDuctAccessoryForm(uiapp, ozkEntities);
            mainForm.Show();

            return Result.Succeeded;
        }
    }
}
