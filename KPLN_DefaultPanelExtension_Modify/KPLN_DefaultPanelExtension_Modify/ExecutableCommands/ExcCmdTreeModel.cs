using Autodesk.Revit.UI;
using KPLN_ExtraFilter.ExternalCommands;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_DefaultPanelExtension_Modify.ExecutableCommands
{
    internal class ExcCmdTreeModel : IExecutableCommand
    {
        internal const string PluginName = SelectionByModelExtCmd.PluginName;

        public Result Execute(UIApplication app)
        {
            try
            {
                // 1) Путь к внешней dll
#if DEBUG
                string extraFilterDll = $"X:\\BIM\\5_Scripts\\Git_Repo_KPLN\\KPLN_ExtraFilter\\bin\\Debug\\{ModuleData.RevitVersion}\\KPLN_ExtraFilter.dll";
#else
                string extraFilterDll = $"Z:\\Отдел BIM\\03_Скрипты\\09_Модули_KPLN_Loader_v.2\\ExtraFilter\\{ModuleData.RevitVersion}\\KPLN_ExtraFilter.dll";
#endif

                // 2) Загружаем сборку
                Assembly asm = Assembly.LoadFrom(extraFilterDll);

                // 3) Берем тип по полному имени (namespace + class)
                Type type = asm.GetType("KPLN_ExtraFilter.ExternalCommands.SelectionByModelExtCmd", true);

                // Создаем экземпляр типа
                object instance = Activator.CreateInstance(type);

                // Определяем метод ExecuteByUIApp
                MethodInfo executeMethod = type.GetMethod("ExecuteByUIApp");

                // Вызываем метод ExecuteByUIApp, передавая аргументы
                if (executeMethod != null)
                    executeMethod.Invoke(instance, new object[] { app, ViewFilterMode.UserSelection });
                else
                    throw new Exception("Ошибка определения метода через рефлексию. Отправь это разработчику\n");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("KPLN: Ошибка", $"Отправь это разработчику:\n{ex.Message}");
            }

            return Result.Succeeded;
        }
    }
}
