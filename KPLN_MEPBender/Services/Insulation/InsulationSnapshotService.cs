using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using KPLN_MEPBender.Services.Parameters;
using System.Collections.Generic;

namespace KPLN_MEPBender.Services.Insulation
{
    internal sealed class InsulationSnapshotService
    {
        private readonly ParameterSnapshotService _parameterSnapshotService;

        public InsulationSnapshotService(ParameterSnapshotService parameterSnapshotService)
        {
            _parameterSnapshotService = parameterSnapshotService;
        }

        public List<InsulationSnapshot> Capture(Document doc, ElementId hostElementId)
        {
            List<InsulationSnapshot> snapshots = new List<InsulationSnapshot>();
            Element hostElement = doc.GetElement(hostElementId);
            bool isPipeHost = hostElement is Pipe;
            bool isDuctHost = hostElement is Duct;

            if (!isPipeHost && !isDuctHost)
                return snapshots;

            foreach (ElementId insulationId in GetInsulationIds(doc, hostElementId))
            {
                Element insulation = doc.GetElement(insulationId);
                PipeInsulation pipeInsulation = insulation as PipeInsulation;
                if (pipeInsulation != null)
                {
                    snapshots.Add(CreateSnapshot(InsulationSnapshotKind.PipeInsulation, pipeInsulation));
                    continue;
                }

                DuctInsulation ductInsulation = insulation as DuctInsulation;
                if (ductInsulation != null)
                    snapshots.Add(CreateSnapshot(InsulationSnapshotKind.DuctInsulation, ductInsulation));
            }

            if (!isDuctHost)
                return snapshots;

            foreach (ElementId liningId in GetLiningIds(doc, hostElementId))
            {
                DuctLining ductLining = doc.GetElement(liningId) as DuctLining;
                if (ductLining != null)
                    snapshots.Add(CreateSnapshot(InsulationSnapshotKind.DuctLining, ductLining));
            }

            return snapshots;
        }

        public List<ElementId> Apply(Document doc, ElementId hostElementId, IEnumerable<InsulationSnapshot> snapshots)
        {
            List<ElementId> createdIds = new List<ElementId>();

            foreach (InsulationSnapshot snapshot in snapshots)
            {
                Element created = Create(doc, hostElementId, snapshot);
                if (created == null)
                    continue;

                _parameterSnapshotService.Apply(created, snapshot.Parameters);
                createdIds.Add(created.Id);
            }

            return createdIds;
        }

        private IEnumerable<ElementId> GetInsulationIds(Document doc, ElementId hostElementId)
        {
            try
            {
                return InsulationLiningBase.GetInsulationIds(doc, hostElementId);
            }
            catch
            {
                return new List<ElementId>();
            }
        }

        private IEnumerable<ElementId> GetLiningIds(Document doc, ElementId hostElementId)
        {
            try
            {
                return InsulationLiningBase.GetLiningIds(doc, hostElementId);
            }
            catch
            {
                return new List<ElementId>();
            }
        }

        private InsulationSnapshot CreateSnapshot(InsulationSnapshotKind kind, InsulationLiningBase insulation)
        {
            return new InsulationSnapshot(
                kind,
                insulation.GetTypeId(),
                insulation.Thickness,
                _parameterSnapshotService.Capture(insulation));
        }

        private Element Create(Document doc, ElementId hostElementId, InsulationSnapshot snapshot)
        {
            try
            {
                switch (snapshot.Kind)
                {
                    case InsulationSnapshotKind.PipeInsulation:
                        return PipeInsulation.Create(doc, hostElementId, snapshot.TypeId, snapshot.Thickness);
                    case InsulationSnapshotKind.DuctInsulation:
                        return DuctInsulation.Create(doc, hostElementId, snapshot.TypeId, snapshot.Thickness);
                    case InsulationSnapshotKind.DuctLining:
                        return DuctLining.Create(doc, hostElementId, snapshot.TypeId, snapshot.Thickness);
                }
            }
            catch
            {
                return null;
            }

            return null;
        }
    }
}
