using Autodesk.Revit.DB;
using System;

namespace KPLN_IOSClasher.Core
{
    internal class FamilyLoadOptions : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            if (familyInUse)
                overwriteParameterValues = true;
            else
                overwriteParameterValues = false;

            return true;
        }

        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
        {
            throw new NotImplementedException();
        }
    }
}
