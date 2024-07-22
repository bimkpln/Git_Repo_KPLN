using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_Loader.Common;
using KPLN_Tools.Common.OVVK_System;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Tools.ExecutableCommand
{
    internal class OZKDuctAccessoryMainCommand : IExecutableCommand
    {
        private readonly OZKDuctAccessoryEntity[] _ozkDuctAccessoryEntities;
        private readonly ExtensibleStorageBuilder _extensibleStorageBuilder;
        private readonly Guid _markParamGuid = new Guid("2204049c-d557-4dfc-8d70-13f19715e46d");
        private readonly Guid _heightParamGuid = new Guid("3bec665b-74a7-4a2a-b696-8ba6eab0b6e8");
        private readonly Guid _widthParamGuid = new Guid("fd1b8e1f-13f5-4aa9-b3c4-d5e75af268d5");
        private readonly Guid _diamParamGuid = new Guid("f27d783d-2505-477e-9292-1b13689e06a4");

        public OZKDuctAccessoryMainCommand(OZKDuctAccessoryEntity[] ozkDuctAccessoryEntities)
        {
            _ozkDuctAccessoryEntities = ozkDuctAccessoryEntities;

            _extensibleStorageBuilder = new ExtensibleStorageBuilder(
                new Guid("85c46d4e-fdc6-424e-909a-27af56597328"),
                "Last_Run",
                "KPLN_OZKAccessory");
        }

        public Result Execute(UIApplication app)
        {
            using (Transaction t = new Transaction(app.ActiveUIDocument.Document, $"KPLN: Клапаны ОЗК"))
            {
                t.Start();

                #region Работа с ExtensibleStorage
                Document doc = app.ActiveUIDocument.Document;
                ProjectInfo pi = doc.ProjectInformation;
                Element piElem = pi as Element;
                _extensibleStorageBuilder.SetStorageData_TimeRunLog(piElem, app.Application.Username, DateTime.Now);
                #endregion

                foreach (OZKDuctAccessoryEntity entity in _ozkDuctAccessoryEntities)
                {
                    if (CheckParams())
                    {
                        string prefParamData = entity.PreffixParam.AsString();
                        string sufParamData = entity.SuffixParam.AsString();
                        foreach (FamilyInstance famInst in entity.CurrentFamilyInstances)
                        {
                            string sizeData = null;
                            Parameter diamParam = famInst.get_Parameter(_diamParamGuid);
                            if (diamParam != null)
                                sizeData = $"ø{diamParam.AsValueString()}";
                            else
                            {
                                Parameter widthParam = famInst.get_Parameter(_widthParamGuid);
                                Parameter heightParam = famInst.get_Parameter(_heightParamGuid);
                                sizeData = $"{widthParam.AsValueString()}x{heightParam.AsValueString()}";
                            }
                            famInst.get_Parameter(_markParamGuid).Set($"{prefParamData}{sizeData}{sufParamData}");
                        }
                    }
                    else
                        return Result.Cancelled;

                }
                
                t.Commit();
                return Result.Succeeded;
            }
        }

        private bool CheckParams()
        {
            IEnumerable<OZKDuctAccessoryEntity> errorColl = _ozkDuctAccessoryEntities
                .Where(ent => 
                    ent.CurrentFamilyInstances.Where(fi => 
                        fi.get_Parameter(_markParamGuid) == null
                        || (fi.get_Parameter(_heightParamGuid) == null && fi.get_Parameter(_widthParamGuid) == null && fi.get_Parameter(_diamParamGuid) == null))
                        .Any()
                    || ent.PreffixParam == null 
                    || ent.SuffixParam == null);
            
            if (errorColl.Any())
            {
                Print($"У элемента с id: {errorColl.FirstOrDefault().CurrentFamilyInstances.FirstOrDefault().Id} нет одного из параметров для анализа: " +
                    $"КП_О_Марка (экземпляр), КП_И_Высота подключения (экземпляр,для прямоугольных), КП_И_Ширина подключения (экземпляр,для прямоугольных), КП_И_Диаметр условный (экземпляр,для круглых), АВВ_Часть марки_До размера (тип), АВВ_Часть марки_После размера (тип). " +
                    $"Работа ЭКСТРЕННО завершена.", 
                    MessageType.Error);
                
                return false;
            }
            
            return true;
        }
    }
}
