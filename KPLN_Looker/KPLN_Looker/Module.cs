using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Loader.Core.SQLiteData;
using KPLN_Looker.Services;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Looker
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private readonly User _user = KPLN_Loader.Application.CurrentRevitUser;

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {

            try
            {
                // Перезапись ini-файла
                INIFileService iNIFileService = new INIFileService(_user, application.ControlledApplication.VersionNumber);
                if (!iNIFileService.OverwriteINIFile())
                {
                    throw new Exception($"Ошибка при перезаписи ini-файла");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Print($"Ошибка: {ex.Message}", MessageType.Error);
                return Result.Failed;
            }
        }
    }
}
