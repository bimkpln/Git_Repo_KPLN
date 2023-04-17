using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandDimensionHelper : IExternalCommand
    {
        private static List<DimensionDTO> _docDimensionsList = new List<DimensionDTO>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            TaskDialog taskDialog = new TaskDialog("Выбери действие");
            taskDialog.MainIcon = TaskDialogIcon.TaskDialogIconInformation;
            taskDialog.MainContent = "Записать значения размеров по выбранной связи - нажми Да.\nВосстановить размеры - нажми Повтор.";
            taskDialog.CommonButtons = TaskDialogCommonButtons.Retry | TaskDialogCommonButtons.Yes;
            TaskDialogResult userInput = taskDialog.Show();
            if (userInput == TaskDialogResult.Cancel)
                return Result.Cancelled;
            else if (userInput == TaskDialogResult.Yes)
            {
                if (_docDimensionsList.Count > 0)
                    _docDimensionsList.Clear();
                
                if (DimensionDataPrepare(doc))
                {
                    Print(
                        $"Данные по размерам внесены в память! Теперь можно подгрузить связь и запустить скрипт на восстановление размеров",
                        KPLN_Loader.Preferences.MessageType.Success);
                    return Result.Succeeded;
                }
                else
                    return Result.Failed;
            }
            else
            {
                DimensionReCreation(doc);
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Получаю размеры из документа
        /// </summary>
        private bool DimensionDataPrepare(Document doc)
        {
            List<DimensionDTO> result = new List<DimensionDTO>();

            var dimFEC = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Dimensions).Cast<Dimension>();
            foreach (Dimension dim in dimFEC)
            {
                View view = dim.View;
                Curve curve = dim.Curve;
                ReferenceArray refArr = dim.References;
                DimensionType type = doc.GetElement(dim.GetTypeId()) as DimensionType;
                if (view != null && !refArr.IsEmpty)
                {
                    XYZ dimPosit = dim.LeaderEndPosition;
                    _docDimensionsList.Add(new DimensionDTO
                    {
                        DimId = dim.Id,
                        DimViewId = view.Id,
                        DimCurve = curve,
                        DimLidEndPostion = dim.LeaderEndPosition,
                        DimRefArray = refArr,
                        DimType = type,
                    });
                }
                else
                {
                    Print(
                        $"У элемента с id: {dim.Id} проблемы: или нет вида, или нет ссылок на элементы для размера",
                        KPLN_Loader.Preferences.MessageType.Warning);
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Создание размеров по данным из xml-файла
        /// </summary>
        private void DimensionReCreation(Document doc)
        {
            using (Transaction t = new Transaction(doc))
            {
                t.Start("KPLN_Восстановить размеры");

                List<Dimension> dimColl = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Dimensions).Cast<Dimension>().ToList();
                var deletedDimColl = _docDimensionsList.Where(x => !dimColl.Select(d => d.Id).Contains(x.DimId));
                foreach (DimensionDTO dimDTO in deletedDimColl)
                {
                    RevitLinkInstance tempLink = null;
                    Reference tempRef = null;
                    foreach(Reference refItem in dimDTO.DimRefArray)
                    {
                        RevitLinkInstance linkInst = doc.GetElement(refItem.ElementId) as RevitLinkInstance;
                        if (linkInst != null)
                        {
                            linkInst.GetTransform();
                            tempLink = linkInst;
                        }
                        else
                        {
                            tempRef = refItem;
                        }
                    }


                    Reference linkRef = new Reference(tempLink);
                    
                    
                    dimDTO.DimRefArray = new ReferenceArray();
                    dimDTO.DimRefArray.Append(tempRef);
                    dimDTO.DimRefArray.Append(linkRef.CreateReferenceInLink());

                    Dimension newDim = doc.Create.NewDimension(
                        doc.GetElement(dimDTO.DimViewId) as View,
                        dimDTO.DimCurve as Line,
                        dimDTO.DimRefArray,
                        dimDTO.DimType);
                    
                    //newDim.LeaderEndPosition = dimDTO.DimLidEndPostion;
                }

                t.Commit();
            }
        }
    }
}
