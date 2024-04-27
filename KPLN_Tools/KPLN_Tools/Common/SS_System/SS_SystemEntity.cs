using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using System;
using System.Collections.Generic;

namespace KPLN_Tools.Common.SS_System
{
    internal class SS_SystemEntity
    {
        public SS_SystemEntity(Element elem)
        {
            CurrentElem = elem;

            CurrentFamInst = (FamilyInstance)CurrentElem;
            if (CurrentFamInst == null)
                throw new Exception($"Не удалось преобразовать в FamilyInstance: {CurrentElem.Id}. Скинь разработчику!");

            CurrentElSystemSet = CurrentFamInst.MEPModel.GetElectricalSystems();
        }

        public Element CurrentElem { get; private set; }

        public FamilyInstance CurrentFamInst { get; private set; }

        public ISet<ElectricalSystem> CurrentElSystemSet { get; set; }


    }
}
