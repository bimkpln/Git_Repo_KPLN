using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace KPLN_ViewsAndLists_Ribbon.Common
{
    internal class UserErrorException : Exception
    {
        public UserErrorException(string errorMessage)
        {
            ErrorMessage = errorMessage;
        }

        public UserErrorException(string errorMessage, List<Element> errorElements) : this(errorMessage)
        {
            ErrorElements = errorElements;
        }

        public List<Element> ErrorElements { get; }

        public string ErrorMessage { get; }
    }
}