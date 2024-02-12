using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace KPLN_Tools.Common
{
    public class MonitorEntity
    {
        internal MonitorEntity(
            Element modelElement,
            Solid modelElemSolid,
            HashSet<Parameter> modelParameters,
            MonitorLinkEntity monitorLinkEntity)
        {
            ModelElement = modelElement;
            ModelElementSolid = modelElemSolid;
            ModelParameters = modelParameters;
            CurrentMonitorLinkEntity = monitorLinkEntity;
        }

        internal Element ModelElement { get; set; }

        internal Solid ModelElementSolid { get; set; }
        internal HashSet<Parameter> ModelParameters { get; set; }

        internal MonitorLinkEntity CurrentMonitorLinkEntity { get; set; }
    }
}