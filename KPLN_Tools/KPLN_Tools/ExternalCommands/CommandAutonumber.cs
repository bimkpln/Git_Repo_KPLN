using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class CommandAutonumber : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string exePath = "\\Source\\autonumber.exe";
            string fullPath = string.Concat(assemblyPath, exePath);
            if (!File.Exists(fullPath))
            {
                message = message + "Не найден исполняемый файл " + fullPath;
                return (Result)1;
            }
            try
            {
                Process.Start(fullPath);
            }
            catch
            {
                message = message + "Не удалось запустить исполняемый файл " + fullPath;
                return (Result)1;
            }
            return (Result)0;
        }
    }
}
