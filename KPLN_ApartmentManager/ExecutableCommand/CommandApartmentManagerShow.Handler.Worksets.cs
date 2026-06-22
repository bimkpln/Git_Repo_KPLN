using Autodesk.Revit.DB;
using KPLN_ApartmentManager.Forms;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ApartmentManager.ExecutableCommand
{
    internal partial class ApartmentManagerExternalHandler
    {
        private class ApartmentWorksetTargets
        {
            public int? WallWorksetId { get; set; }
            public int? DoorWorksetId { get; set; }
            public int? RoomWorksetId { get; set; }
            public int? FurnitureWorksetId { get; set; }
            public int? PlumbingWorksetId { get; set; }
            public int? WindowWorksetId { get; set; }
        }

        private static List<string> BuildUserWorksetOptions(Document doc)
        {
            List<string> result = new List<string>();
            if (doc == null)
                return result;

            try
            {
                if (!doc.IsWorkshared)
                    return result;

                foreach (Workset workset in new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset))
                {
                    if (workset != null && !string.IsNullOrWhiteSpace(workset.Name))
                        result.Add(workset.Name);
                }
            }
            catch
            {
            }

            return result
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();
        }

        private static ApartmentWorksetTargets ResolveApartmentWorksetTargets(Document doc, ApartmentPresetData preset)
        {
            return new ApartmentWorksetTargets
            {
                WallWorksetId = ResolveUserWorksetId(doc, preset != null ? preset.WallWorksetName : null),
                DoorWorksetId = ResolveUserWorksetId(doc, preset != null ? preset.DoorWorksetName : null),
                RoomWorksetId = ResolveUserWorksetId(doc, preset != null ? preset.RoomWorksetName : null),
                FurnitureWorksetId = ResolveUserWorksetId(doc, preset != null ? preset.FurnitureWorksetName : null),
                PlumbingWorksetId = ResolveUserWorksetId(doc, preset != null ? preset.PlumbingWorksetName : null),
                WindowWorksetId = ResolveUserWorksetId(doc, preset != null ? preset.WindowWorksetName : null)
            };
        }

        private static int? ResolveUserWorksetId(Document doc, string worksetName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(worksetName) ||
                string.Equals(worksetName, ApartmentPresetData.NoWorksetSelection, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            try
            {
                if (!doc.IsWorkshared)
                    return null;

                Workset workset = new FilteredWorksetCollector(doc)
                    .OfKind(WorksetKind.UserWorkset)
                    .FirstOrDefault(x => x != null && string.Equals(x.Name, worksetName, StringComparison.OrdinalIgnoreCase));

                return workset != null
                    ? (int?)workset.Id.IntegerValue
                    : null;
            }
            catch
            {
                return null;
            }
        }

        private static bool TryAssignElementToWorkset(Element element, int? worksetId)
        {
            if (element == null || !worksetId.HasValue)
                return false;

            try
            {
                Parameter parameter = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                if (parameter == null || parameter.IsReadOnly || parameter.StorageType != StorageType.Integer)
                    return false;

                parameter.Set(worksetId.Value);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
