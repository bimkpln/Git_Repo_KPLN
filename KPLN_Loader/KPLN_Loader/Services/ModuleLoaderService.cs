using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Loader.Core.Entities;
using NLog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace KPLN_Loader.Services
{
    internal sealed class ModuleLoaderService
    {
        private static readonly string[] DllSeparator = { ".dll" };

        private readonly Logger _logger;
        private readonly string _ribbonName;
        private readonly List<IExternalModule> _moduleInstances;
        private readonly Action<string, bool> _statusCallback;

        internal ModuleLoaderService(Logger logger, string ribbonName, List<IExternalModule> moduleInstances, Action<string, bool> statusCallback)
        {
            _logger = logger;
            _ribbonName = ribbonName;
            _moduleInstances = moduleInstances;
            _statusCallback = statusCallback;
        }

        internal int LoadLocalModules(UIControlledApplication application, IEnumerable<Module> modules, Func<Module, DirectoryInfo> moduleDirectoryProvider)
        {
            int uploadModules = 0;

            foreach (Module module in modules)
            {
                if (module == null)
                {
                    PublishStatus(string.Format(LoaderMessageService.ModuleMessages.MissingModule, "<null>"), true);
                    continue;
                }

                bool isModuleLoad = false;
                try
                {
                    DirectoryInfo targetDirInfo = moduleDirectoryProvider(module);
                    if (targetDirInfo == null)
                    {
                        PublishStatus(string.Format(LoaderMessageService.ModuleMessages.CopyFailed, module.Name), true);
                        continue;
                    }

                    isModuleLoad = TryLoadFromDirectory(application, targetDirInfo, module.Name, module.IsLibraryModule, out string moduleVersion);
                    if (isModuleLoad)
                    {
                        uploadModules++;
                        string msg = module.IsLibraryModule
                            ? string.Format(LoaderMessageService.ModuleMessages.LibraryLoaded, module.Name, moduleVersion, "скопирован")
                            : string.Format(LoaderMessageService.ModuleMessages.ModuleLoaded, module.Name, moduleVersion);

                        PublishStatus(msg, false);
                    }
                    else
                    {
                        PublishStatus(string.Format(LoaderMessageService.ModuleMessages.MissingDll, module.Name), true);
                    }
                }
                catch (Exception ex)
                {
                    string msg = string.Format(LoaderMessageService.ModuleMessages.LocalLoadError, module.Name);
                    _logger.Error(msg + $" \n{ex}");
                    _statusCallback?.Invoke(msg, true);
                }
            }

            return uploadModules;
        }

        internal int LoadExtraNetModules(UIControlledApplication application, DirectoryInfo modulesDirectory)
        {
            int uploadModules = 0;

            foreach (DirectoryInfo dir in modulesDirectory.GetDirectories())
            {
                bool isLibrary = dir.Name.Contains("Library");
                bool isModuleLoad = TryLoadFromDirectory(application, dir, dir.Name, isLibrary, out string moduleVersion);

                if (isModuleLoad)
                {
                    uploadModules++;
                    string msg = isLibrary
                        ? string.Format(LoaderMessageService.ModuleMessages.LibraryLoaded, dir.Name, moduleVersion, "активирован")
                        : string.Format(LoaderMessageService.ModuleMessages.ModuleLoaded, dir.Name, moduleVersion);

                    PublishStatus(msg, false);
                }
                else
                {
                    PublishStatus(string.Format(LoaderMessageService.ModuleMessages.MissingDll, dir.Name), true);
                }
            }

            return uploadModules;
        }

        private bool TryLoadFromDirectory(UIControlledApplication application, DirectoryInfo sourceDirectory, string moduleName, bool isLibrary, out string moduleVersion)
        {
            moduleVersion = "-";
            bool isModuleLoad = false;

            foreach (FileInfo file in sourceDirectory.GetFiles())
            {
                if (!CheckModuleName(file.Name))
                    continue;

                var moduleAssembly = System.Reflection.Assembly.LoadFrom(file.FullName);
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(moduleAssembly.Location);
                if (moduleAssembly.FullName.Contains("KPLN"))
                    moduleVersion = fvi.FileVersion;

                if (isLibrary)
                {
                    isModuleLoad = true;
                    continue;
                }

                Type implementationType = moduleAssembly.GetType(file.Name.Split(DllSeparator, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() + ".Module", false);
                if (implementationType == null)
                    continue;

                IExternalModule moduleInstance = Activator.CreateInstance(implementationType) as IExternalModule;
                Result loadingResult = moduleInstance.Execute(application, _ribbonName);
                if (loadingResult == Result.Succeeded)
                {
                    _moduleInstances.Add(moduleInstance);
                    isModuleLoad = true;
                }
                else
                {
                    string msg = string.Format(LoaderMessageService.ModuleMessages.ModuleNotActivated, moduleName);
                    _logger.Error(msg + $" Проблема: \n{loadingResult}");
                    _statusCallback?.Invoke(msg + " Подробнее - см. файл логов", true);
                }
            }

            return isModuleLoad;
        }

        private bool CheckModuleName(string fileName) => fileName.Split(DllSeparator, StringSplitOptions.None).Length > 1 && !fileName.Contains(".config");

        private void PublishStatus(string message, bool isError)
        {
            if (isError)
                _logger.Error(message);
            else
                _logger.Info(message);

            _statusCallback?.Invoke(message, isError);
        }
    }
}
