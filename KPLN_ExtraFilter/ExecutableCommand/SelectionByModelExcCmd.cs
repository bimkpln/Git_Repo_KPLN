using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Common;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ExtraFilter.ExecutableCommand
{
    internal class SelectionByModelExcCmd : IExecutableCommand
    {
        private readonly SelectionByModelM _entity;

        public SelectionByModelExcCmd(SelectionByModelM entity)
        {
            _entity = entity;
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uiDoc = app.ActiveUIDocument;
            if (uiDoc == null)
                return Result.Cancelled;

            Document doc = uiDoc.Document;
            try
            {
                ElementId[] resultIDToSelect = null;

                List<Element> userSelFromFormElems = new List<Element>();
                TreeElementEntity.GetAllCheckedRevitElemsFromTreeElemColl(_entity.TreeElemEntities, ref userSelFromFormElems);
                if (userSelFromFormElems.Count() == 0)
                    return Result.Cancelled;

                
                IEnumerable<ElementId> userSelFromFormIDs = userSelFromFormElems.Select(el => el.Id);
                switch (_entity.How_SelectFilterMode)
                {
                    case SelectFilterMode.CreateNew:
                        resultIDToSelect = userSelFromFormIDs.ToArray();
                        break;
                    case SelectFilterMode.AddCurrent:
                        IEnumerable<ElementId> userSelFromDocIDs = _entity.UserSelElems.Select(el => el.Id);
                        resultIDToSelect = userSelFromDocIDs.Union(userSelFromFormIDs).ToArray();
                        break;
                }
                
                if (resultIDToSelect != null)
                    uiDoc.Selection.SetElementIds(resultIDToSelect);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                HtmlOutput.Print($"Ошибка попытки выбора подобных. Отправь разработчику: {ex.Message}",
                    MessageType.Error);

                return Result.Cancelled;
            }
        }
    }
}
