using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_MEPBender.Services.Routing
{
    internal static class MepCurveConnectorUtils
    {
        public static Connector GetClosestEndConnector(MEPCurve curve, XYZ point)
        {
            if (curve?.ConnectorManager == null || point == null)
                return null;

            Connector closest = null;
            double minDistance = double.MaxValue;

            foreach (Connector connector in curve.ConnectorManager.Connectors)
            {
                if (connector.ConnectorType != ConnectorType.End)
                    continue;

                double distance = connector.Origin.DistanceTo(point);
                if (distance < minDistance)
                {
                    minDistance = distance;
                    closest = connector;
                }
            }

            return closest;
        }

        public static IEnumerable<ExternalConnectorInfo> GetExternalConnections(MEPCurve source, XYZ start, XYZ end)
        {
            if (source?.ConnectorManager == null)
                yield break;

            foreach (Connector connector in source.ConnectorManager.Connectors)
            {
                if (connector.ConnectorType != ConnectorType.End)
                    continue;

                XYZ sourceEndpoint = connector.Origin.DistanceTo(start) <= connector.Origin.DistanceTo(end)
                    ? start
                    : end;

                foreach (Connector connectedConnector in connector.AllRefs)
                {
                    if (connectedConnector.Owner == null || connectedConnector.Owner.Id == source.Id)
                        continue;

                    yield return new ExternalConnectorInfo(connectedConnector.Owner.Id, sourceEndpoint);
                }
            }
        }

        public static Connector GetClosestConnector(Element element, XYZ point)
        {
            IEnumerable<Connector> connectors = GetConnectors(element);
            if (connectors == null)
                return null;

            return connectors
                .Where(c => c.ConnectorType == ConnectorType.End)
                .OrderBy(c => c.Origin.DistanceTo(point))
                .FirstOrDefault();
        }

        public static bool TryConnect(Document doc, Connector firstConnector, Connector secondConnector, bool useElbow, out Element createdElement)
        {
            createdElement = null;

            if (firstConnector == null || secondConnector == null)
                return false;

            try
            {
                if (!firstConnector.IsConnectedTo(secondConnector))
                {
                    if (useElbow)
                        createdElement = doc.Create.NewElbowFitting(firstConnector, secondConnector);
                    else
                        firstConnector.ConnectTo(secondConnector);
                }

                return true;
            }
            catch
            {
                try
                {
                    if (!firstConnector.IsConnectedTo(secondConnector))
                        firstConnector.ConnectTo(secondConnector);

                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }

        private static IEnumerable<Connector> GetConnectors(Element element)
        {
            MEPCurve mepCurve = element as MEPCurve;
            if (mepCurve?.ConnectorManager != null)
            {
                foreach (Connector connector in mepCurve.ConnectorManager.Connectors)
                    yield return connector;
            }

            FamilyInstance familyInstance = element as FamilyInstance;
            if (familyInstance?.MEPModel?.ConnectorManager == null)
                yield break;

            foreach (Connector connector in familyInstance.MEPModel.ConnectorManager.Connectors)
                yield return connector;
        }
    }
}
