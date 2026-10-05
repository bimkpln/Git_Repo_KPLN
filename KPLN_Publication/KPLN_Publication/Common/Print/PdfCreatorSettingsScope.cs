using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Publication
{
    /// <summary>Временный профиль автосохранения; пользовательские значения восстанавливаются точно.</summary>
    internal sealed class PdfCreatorSettingsScope : IDisposable
    {
        private const string Root = @"Software\pdfforge\PDFCreator\Settings";
        private const string Profile = Root + @"\ConversionProfiles\0";
        private readonly List<Action> _restore = new List<Action>();

        internal static void Validate()
        {
            using (var profile = Registry.CurrentUser.OpenSubKey(Profile))
            using (var mappings = Registry.CurrentUser.OpenSubKey(Root + @"\ApplicationSettings\PrinterMappings"))
            {
                if (profile == null || mappings == null)
                    throw new InvalidOperationException("PDFCreator не настроен: отсутствует профиль или привязка принтера.");

                string guid = profile.GetValue("Guid") as string;
                bool matched = false;
                foreach (string name in mappings.GetSubKeyNames())
                {
                    using (var mapping = mappings.OpenSubKey(name))
                    {
                        if ((string)mapping.GetValue("PrinterName") != "PDFCreator")
                            continue;

                        if (string.IsNullOrEmpty(guid) || (string)mapping.GetValue("ProfileGuid") != guid)
                            throw new InvalidOperationException("PDFCreator привязан не к профилю 0; печать не запускалась.");

                        matched = true;
                    }
                }
                if (!matched)
                    throw new InvalidOperationException("Не найдена явная привязка принтера PDFCreator к профилю 0.");
            }
        }

        internal PdfCreatorSettingsScope(string folder, bool useOrientation, bool portrait)
        {
            Validate();
            try
            {
                // Исключаем действия профиля: просмотрщик, письма, загрузки, повторную печать и скрипты.
                using (var profile = Registry.CurrentUser.OpenSubKey(Profile))
                {
                    foreach (string child in profile.GetSubKeyNames())
                    {
                        using (var action = profile.OpenSubKey(child))
                        {
                            if (child != "AutoSave" && action.GetValueNames().Contains("Enabled"))
                                Set(Profile + "\\" + child, "Enabled", "False");
                        }
                    }
                }

                Set(Profile + @"\AutoSave", "Enabled", "True");
                Set(Profile + @"\AutoSave", "EnsureUniqueFilenames", "True");
                Set(Profile, "FileNameTemplate", "<InputFilename>");
                // PDFCreator декодирует строки реестра: одиночный обратный слеш теряет путь.
                // Как в штатной пакетной печати, экранируем только новое значение, не резервную копию.
                Set(Profile, "TargetDirectory", EncodeDirectory(folder));
                Set(Profile, "OutputFormat", "Pdf");
                Set(Profile, "OpenViewer", "False");
                Set(Profile, "OpenWithPdfArchitect", "False");
                Set(Profile + @"\PdfSettings", "PageOrientation", useOrientation ? (portrait ? "Portrait" : "Landscape") : "Automatic");
            }
            catch
            {
                Dispose();
                throw;
            }
        }

        internal static string EncodeDirectory(string folder)
        {
            return folder.Replace("\\", "\\\\");
        }

        private void Set(string path, string name, string value)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(path, true))
            {
                if (key == null)
                    throw new InvalidOperationException("Отсутствует раздел настроек PDFCreator: " + path);

                bool existed = key.GetValueNames().Contains(name);
                object previous = key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                RegistryValueKind kind = existed ? key.GetValueKind(name) : RegistryValueKind.String;
                _restore.Add(() =>
                {
                    using (var target = Registry.CurrentUser.OpenSubKey(path, true))
                    {
                        if (target == null)
                            throw new InvalidOperationException("Исчез раздел настроек PDFCreator: " + path);
                        if (existed)
                            target.SetValue(name, previous, kind);
                        else
                            target.DeleteValue(name, false);
                    }
                });
                key.SetValue(name, value, RegistryValueKind.String);
            }
        }

        public void Dispose()
        {
            var errors = new List<Exception>();
            for (int i = _restore.Count - 1; i >= 0; i--)
            {
                try { _restore[i](); }
                catch (Exception ex) { errors.Add(ex); }
            }
            _restore.Clear();
            if (errors.Count != 0)
                throw new AggregateException("Не удалось полностью восстановить настройки PDFCreator.", errors);
        }
    }
}
