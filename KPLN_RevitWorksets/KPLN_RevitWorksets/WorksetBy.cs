#region License
/*Данный код опубликован под лицензией Creative Commons Attribution-ShareAlike.
Разрешено использовать, распространять, изменять и брать данный код за основу для производных в коммерческих и
некоммерческих целях, при условии указания авторства и если производные лицензируются на тех же условиях.
Код поставляется "как есть". Автор не несет ответственности за возможные последствия использования.
Зуев Александр, 2020, все права защищены.
This code is listed under the Creative Commons Attribution-ShareAlike license.
You may use, redistribute, remix, tweak, and build upon this work non-commercially and commercially,
as long as you credit the author by linking back and license your new creations under the same terms.
This code is provided 'as is'. Author disclaims any implied warranty.
Zuev Aleksandr, 2020, all rigths reserved.*/
#endregion
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace KPLN_RevitWorksets
{
    public abstract class WorksetBy
    {
        public string WorksetName;

        public Workset GetWorkset(Document doc)
        {
            IList<Workset> userWorksets = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets();

            bool checkNotExists = WorksetTable.IsWorksetNameUnique(doc, WorksetName);

            if (!checkNotExists)
            {
                Workset wset = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets()
                .Where(w => w.Name == WorksetName)
                .First();
                return wset;
            }
            else
            {
                Workset wset = Workset.Create(doc, WorksetName);
                return wset;
            }
        }


        public static void SetWorkset(Element elem, Workset w)
        {
            Parameter wsparam = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
            if (wsparam == null) return;
            if (wsparam.IsReadOnly) return;

            bool elemNonGroup = (elem.GroupId == null) || (elem.GroupId == ElementId.InvalidElementId);

            if (elemNonGroup)
            {
                wsparam.Set(w.Id.IntegerValue);
            }
            else
            {
                Group gr = elem.Document.GetElement(elem.GroupId) as Group;
                SetWorkset(gr, w);
            }
        }
    }
}
