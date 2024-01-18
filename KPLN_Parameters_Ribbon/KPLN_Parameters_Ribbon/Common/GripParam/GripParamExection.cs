using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace KPLN_Parameters_Ribbon.Common.GripParam
{
    internal class GripParamExection : Exception
    {
        public GripParamExection(string errorMessage)
        {
            ErrorMessage = errorMessage;
        }

        public GripParamExection(string errorMessage, List<Element> errorElements) : this(errorMessage)
        {
            ErrorElements = errorElements;
        }

        public List<Element> ErrorElements { get; }

        public string ErrorMessage { get; }
    }
}
