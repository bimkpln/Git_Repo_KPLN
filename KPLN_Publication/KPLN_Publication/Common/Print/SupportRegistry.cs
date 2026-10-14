using System;
using System.IO;
using Microsoft.Win32;

namespace KPLN_Publication
{
    public enum OrientationType { Portrait, Landscape, Automatic };

    public static class SupportRegistry
    {
        private const string SettingsRoot = @"Software\pdfforge\PDFCreator\Settings";

        internal static string GetPrinterProfilePath()
        {
            using (RegistryKey mappings = Registry.CurrentUser.OpenSubKey(SettingsRoot + @"\ApplicationSettings\PrinterMappings"))
            using (RegistryKey profiles = Registry.CurrentUser.OpenSubKey(SettingsRoot + @"\ConversionProfiles"))
            {
                if (mappings == null || profiles == null)
                    throw new InvalidOperationException("PDFCreator не настроен: отсутствуют профили или привязки принтеров.");

                string profileGuid = null;
                foreach (string name in mappings.GetSubKeyNames())
                {
                    using (RegistryKey mapping = mappings.OpenSubKey(name))
                    {
                        if (!string.Equals(mapping?.GetValue("PrinterName") as string, "PDFCreator", StringComparison.OrdinalIgnoreCase))
                            continue;

                        string mappedGuid = mapping.GetValue("ProfileGuid") as string;
                        if (profileGuid != null && !string.Equals(profileGuid, mappedGuid, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("У принтера PDFCreator неоднозначная привязка профиля.");
                        profileGuid = mappedGuid;
                    }
                }

                if (string.IsNullOrWhiteSpace(profileGuid))
                    throw new InvalidOperationException("Назначьте принтеру PDFCreator профиль в настройках PDFCreator.");

                string profilePath = null;
                foreach (string name in profiles.GetSubKeyNames())
                {
                    using (RegistryKey profile = profiles.OpenSubKey(name))
                    {
                        if (!string.Equals(profile?.GetValue("Guid") as string, profileGuid, StringComparison.OrdinalIgnoreCase))
                            continue;

                        if (profilePath != null)
                            throw new InvalidOperationException("Несколько профилей PDFCreator имеют одинаковый GUID.");
                        profilePath = SettingsRoot + @"\ConversionProfiles\" + name;
                    }
                }

                return profilePath ?? throw new InvalidOperationException("Профиль, назначенный принтеру PDFCreator, не найден.");
            }
        }

        internal static string EncodeDirectory(string folder)
        {
            // PDFCreator декодирует строки: одиночные слеши превращают путь в пустую строку.
            return folder.Replace("\\", "\\\\");
        }

        public static void ActivateSettingsForPDFCreator(string outputFolder)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(outputFolder) || !Directory.Exists(outputFolder))
                    throw new DirectoryNotFoundException("Не найдена папка экспорта PDF: " + outputFolder);

                // Меняем существующий профиль, назначенный принтеру, без создания новых профилей.
                using (RegistryKey profile = Registry.CurrentUser.OpenSubKey(GetPrinterProfilePath(), true))
                {
                    if (profile == null)
                        throw new InvalidOperationException("Профиль PDFCreator недоступен для записи.");

                    profile.SetValue("FileNameTemplate", "<InputFilename>");
                    profile.SetValue("TargetDirectory", EncodeDirectory(Path.GetFullPath(outputFolder)));
                    profile.SetValue("SaveFileTemporary", "False");
                    profile.SetValue("OutputFormat", "Pdf");

                    using (RegistryKey autoSave = profile.CreateSubKey("AutoSave"))
                        autoSave.SetValue("Enabled", "True");

                    // В PDFCreator 4 действие просмотра хранится в отдельном разделе.
                    using (RegistryKey viewer = profile.CreateSubKey("OpenViewer"))
                    {
                        viewer.SetValue("Enabled", "False");
                        viewer.SetValue("OpenWithPdfArchitect", "False");
                    }

                    // Совместимость со старыми версиями PDFCreator.
                    profile.SetValue("OpenViewer", "False");
                    profile.SetValue("OpenWithPdfArchitect", "False");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Не удалось подготовить PDFCreator к пакетной печати. " + ex.Message, ex);
            }

        }

        public static void SetOrientationForPdfCreator(OrientationType orientType)
        {
            string keyName = "HKEY_CURRENT_USER\\" + GetPrinterProfilePath() + @"\PdfSettings";
            string keyTitle = "PageOrientation";
            string text = Enum.GetName(typeof(OrientationType), orientType);
            string check = "";
            try
            {
                check = Microsoft.Win32.Registry.GetValue(keyName, keyTitle, "default") as string;
                if (check == null || check == "default" || string.IsNullOrEmpty(check))
                {
                    return;
                }
                else
                {
                    if (check != text)
                    {
                        Microsoft.Win32.Registry.SetValue(keyName, keyTitle, text);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = "Не удалось настроить PDFCreator! Если это первый запуск - попробуйте отправить на PDFCreator любой другой документ, например, из Word. После этого запустите печать еще раз.";
                msg += "Устанавливаемое значение ключа реестра: " + text + ", значение ключа реестра: " + check + ". ";
                msg = msg + "Текст ошибки: " + ex.Message;
                MessageBox msgbox = new MessageBox(msg);
                msgbox.ShowDialog();
                throw new Exception(msg);
            }
        }

        public static void RestoreSettingsForPDFCreator()
        {
            try
            {
                Microsoft.Win32.Registry.SetValue(
                    "HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0\\AutoSave", "Enabled", "False");
                Microsoft.Win32.Registry.SetValue(
                    "HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0", "FileNameTemplate", "<Title>");
                Microsoft.Win32.Registry.SetValue(
                    "HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0", "OpenViewer", "True");
                Microsoft.Win32.Registry.SetValue(
                    "HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0", "OpenWithPdfArchitect", "False");
                Microsoft.Win32.Registry.SetValue(
                    "HKEY_CURRENT_USER\\Software\\pdfforge\\PDFCreator\\Settings\\ConversionProfiles\\0", "TargetDirectory", "");
            }
            catch
            {
                string msg = "Не удалось восстановить настройки PDFCreator по умолчанию. Попробуйте переустановить принтер.";
                MessageBox msgbox = new MessageBox(msg);
                msgbox.ShowDialog();
                throw new Exception(msg);
            }
        }
    }
}
