using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace KPLN_Publication
{
    /// <summary>Временный профиль автосохранения; пользовательские значения восстанавливаются точно.</summary>
    internal sealed class PdfCreatorSettingsScope : IDisposable
    {
        private const string Root = @"Software\pdfforge\PDFCreator\Settings";
        private const string Profile = Root + @"\ConversionProfiles\0";
        private readonly List<Action> _restore = new List<Action>();
        private readonly List<Func<bool>> _verify = new List<Func<bool>>();

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
                if (existed && previous == null)
                    throw new InvalidOperationException("Настройки PDFCreator изменились во время чтения: " + path + "\\" + name);

                _restore.Add(() =>
                {
                    using (var target = existed ? Registry.CurrentUser.CreateSubKey(path) : Registry.CurrentUser.OpenSubKey(path, true))
                    {
                        if (existed)
                            target.SetValue(name, previous, kind);
                        else if (target != null)
                            target.DeleteValue(name, false);
                    }
                });
                _verify.Add(() =>
                {
                    using (var target = Registry.CurrentUser.OpenSubKey(path))
                    {
                        bool present = target != null && target.GetValueNames().Contains(name);
                        if (present != existed)
                            return false;
                        if (!existed)
                            return true;

                        object actual = target.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                        if (target.GetValueKind(name) != kind)
                            return false;
                        if (previous is Array && actual is Array)
                            return ((Array)previous).Cast<object>().SequenceEqual(((Array)actual).Cast<object>());

                        return Equals(previous, actual);
                    }
                });
                key.SetValue(name, value, RegistryValueKind.String);
            }
        }

        public void Dispose()
        {
            if (_restore.Count == 0)
                return;

            try
            {
                RestoreAndVerify(_restore, _verify, Thread.Sleep, 8);
            }
            finally
            {
                _restore.Clear();
                _verify.Clear();
            }
        }

        // PDFCreator может пересоздавать разделы после закрытия выходного PDF.
        // Повторяем только восстановление снимка, не отправку задания печати.
        internal static void RestoreAndVerify(IList<Action> restore, IList<Func<bool>> verify, Action<int> wait, int attempts)
        {
            if (restore.Count != verify.Count || attempts < 1)
                throw new ArgumentException("Неверный пакет восстановления настроек.");

            var errors = new List<Exception>();
            wait(250);
            for (int attempt = 0; attempt < attempts; attempt++)
            {
                errors.Clear();
                for (int i = restore.Count - 1; i >= 0; i--)
                {
                    try { restore[i](); }
                    catch (Exception ex) { errors.Add(ex); }
                }

                // Две проверки замечают позднюю запись кэша PDFCreator поверх снимка.
                bool restored = errors.Count == 0;
                for (int check = 0; check < 2; check++)
                {
                    wait(150);
                    foreach (var matches in verify)
                    {
                        try { restored &= matches(); }
                        catch (Exception ex) { errors.Add(ex); restored = false; }
                    }
                }
                if (restored)
                    return;
            }

            if (errors.Count == 0)
                errors.Add(new InvalidOperationException("Значения профиля продолжают изменяться после восстановления."));

            string detail = string.Join("; ", errors.Select(e => e.GetType().Name + ": " + e.Message).Distinct());
            throw new AggregateException("Не удалось восстановить и проверить настройки PDFCreator. " + detail, errors);
        }
    }
}
