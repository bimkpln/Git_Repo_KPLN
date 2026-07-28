using Autodesk.Revit.UI;
using KPLN_CommandsWheel.Forms;
using KPLN_CommandsWheel.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_CommandsWheel.Services
{
    internal static class CommandWindowService
    {
        private static RevitCommandExecutor _executor;

        internal static bool ShowCommandSearch(UIApplication uiapp)
        {
            if (CommandSearchWindow.TryActivateExisting())
            {
                return true;
            }

            UserSettings settings = UserSettingsService.Load();
            List<RevitCommandInfo> commands = RibbonCommandCollector.Collect();
            SelectionCustomCommandService.AddCommands(commands);

            if (commands.Count == 0)
            {
                TaskDialog.Show("KPLN. Штурвал команд. Команды", "Не удалось прочитать команды ленты Revit в текущей сессии.");
                return false;
            }

            bool showLegacyHotkeyNotice = TryMigrateLegacyHotkeys(
                settings,
                commands);
            if (showLegacyHotkeyNotice)
            {
                string notice = "Механизм горячих клавиш изменён. Назначьте клавиши один раз в настройках и перезагрузите Revit.";
                if (!string.IsNullOrWhiteSpace(settings.KeyboardShortcutMigrationMessage))
                {
                    notice += "\n\nПричина: " + settings.KeyboardShortcutMigrationMessage;
                }

                TaskDialog.Show(
                    "KPLN. Штурвал команд. Горячие клавиши",
                    notice);
            }

            CommandSearchWindow window = new CommandSearchWindow(commands, settings, GetExecutor());
            WindowOwnerHelper.Apply(window);
            window.Show();

            return true;
        }

        internal static bool ShowCommandsWheel(UIApplication uiapp)
        {
            if (CommandsWheelWindow.TryActivateExisting())
            {
                return true;
            }

            UserSettings settings = UserSettingsService.Load();
            List<RevitCommandInfo> allCommands = RibbonCommandCollector.Collect();
            SelectionCustomCommandService.AddCommands(allCommands);

            Dictionary<string, RevitCommandInfo> commandsById = allCommands
                .GroupBy(command => command.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            List<RevitCommandInfo> wheelCommands = new List<RevitCommandInfo>();
            foreach (string id in settings.WheelCommandIds.Take(8))
            {
                RevitCommandInfo command;
                if (!string.IsNullOrWhiteSpace(id) && commandsById.TryGetValue(id, out command))
                {
                    wheelCommands.Add(command);
                }
            }

            if (wheelCommands.Count == 0)
            {
                TaskDialog.Show("KPLN. Штурвал команд", "В штурвале нет доступных команд. Добавьте команды через окно \"Команды\".");
                return false;
            }

            CommandsWheelWindow window = new CommandsWheelWindow(wheelCommands, GetExecutor(), settings);
            WindowOwnerHelper.Apply(window);
            WindowPositionHelper.ShowCenteredOnCursor(window);
            window.Show();

            return true;
        }

        internal static void Shutdown()
        {
            RevitCommandExecutor executor = _executor;
            _executor = null;

            if (executor != null)
            {
                try
                {
                    executor.Dispose();
                }
                catch
                {
                    // Revit may already be releasing API resources during shutdown.
                }
                finally
                {
                    RibbonCommandCollector.ClearCache();
                }

                return;
            }

            RibbonCommandCollector.ClearCache();
        }

        private static RevitCommandExecutor GetExecutor()
        {
            if (_executor == null)
            {
                _executor = new RevitCommandExecutor();
            }

            return _executor;
        }

        private static bool TryMigrateLegacyHotkeys(
            UserSettings settings,
            IEnumerable<RevitCommandInfo> commands)
        {
            if (settings == null)
            {
                return false;
            }

            string migrationStatus =
                (settings.KeyboardShortcutMigrationStatus ?? string.Empty).Trim();
            bool migrationIsPending = !settings.LegacyHotkeyMigrationAttempted
                && (migrationStatus.Length == 0
                    || string.Equals(
                        migrationStatus,
                        "Pending",
                        StringComparison.OrdinalIgnoreCase));
            if (!migrationIsPending)
            {
                bool failedOrInterrupted =
                    string.Equals(
                        migrationStatus,
                        "Failed",
                        StringComparison.OrdinalIgnoreCase)
                    || string.Equals(
                        migrationStatus,
                        "Checking",
                        StringComparison.OrdinalIgnoreCase);
                if (!failedOrInterrupted
                    || settings.LegacyHotkeyMigrationNoticeShown)
                {
                    return false;
                }

                settings.LegacyHotkeyMigrationNoticeShown = true;
                settings.KeyboardShortcutMigrationStatus = "Failed";
                if (string.IsNullOrWhiteSpace(settings.KeyboardShortcutMigrationMessage))
                {
                    settings.KeyboardShortcutMigrationMessage =
                        "Предыдущая проверка старых настроек не была завершена.";
                }
                try
                {
                    UserSettingsService.Save(settings);
                }
                catch
                {
                    // Do not block the command window on an informational flag.
                }

                return true;
            }

            bool showNotice = false;
            try
            {
                settings.LegacyHotkeyMigrationAttempted = true;
                settings.LegacyHotkeyMigrationNoticeShown = false;
                settings.KeyboardShortcutMigrationStatus = "Checking";
                settings.KeyboardShortcutMigrationMessage = null;
                UserSettingsService.Save(settings);

                bool hasLegacyHotkeys = HasLegacyHotkey(settings.CommandsWheelHotkey)
                    || HasLegacyHotkey(settings.CommandSearchHotkey);

                if (!hasLegacyHotkeys
                    && (settings.AreKeyboardShortcutsConfigured
                        || !string.IsNullOrWhiteSpace(settings.WheelShortcut)
                        || !string.IsNullOrWhiteSpace(settings.CommandSearchShortcut)))
                {
                    MarkMigrationCompleted(
                        settings,
                        "AlreadyConfigured",
                        "Новые горячие клавиши уже были настроены.");
                    ClearLegacyHotkeys(settings);
                    UserSettingsService.Save(settings);
                    return false;
                }

                string wheelCommandId = KeyboardShortcutService.FindCommandId(
                    commands,
                    typeof(ExternalCommands.CommandsWheel));
                string searchCommandId = KeyboardShortcutService.FindCommandId(
                    commands,
                    typeof(ExternalCommands.CommandSearch));
                if (!hasLegacyHotkeys)
                {
                    string currentWheelShortcut;
                    string currentSearchShortcut;
                    if (KeyboardShortcutService.TryReadCurrentShortcuts(
                        ModuleData.RevitVersion,
                        ModuleData.RevitVersionName,
                        wheelCommandId,
                        searchCommandId,
                        out currentWheelShortcut,
                        out currentSearchShortcut)
                        && (!string.IsNullOrWhiteSpace(currentWheelShortcut)
                            || !string.IsNullOrWhiteSpace(currentSearchShortcut)))
                    {
                        settings.WheelShortcut = currentWheelShortcut ?? string.Empty;
                        settings.CommandSearchShortcut = currentSearchShortcut ?? string.Empty;
                        settings.AreKeyboardShortcutsConfigured = true;
                        MarkMigrationCompleted(
                            settings,
                            "LoadedFromKeyboardShortcutsXml",
                            "Назначения прочитаны из KeyboardShortcuts.xml.");
                        ClearLegacyHotkeys(settings);
                        UserSettingsService.Save(settings);
                        return false;
                    }

                    MarkMigrationCompleted(
                        settings,
                        "NoLegacySettings",
                        "Старые назначения клавиш не найдены.");
                    ClearLegacyHotkeys(settings);
                    UserSettingsService.Save(settings);
                    return false;
                }

                string wheelShortcut;
                string searchShortcut;
                string wheelError;
                string searchError;
                bool wheelIsValid = KeyboardShortcutService.TryConvertLegacyGesture(
                    settings.CommandsWheelHotkey,
                    out wheelShortcut,
                    out wheelError);
                bool searchIsValid = KeyboardShortcutService.TryConvertLegacyGesture(
                    settings.CommandSearchHotkey,
                    out searchShortcut,
                    out searchError);

                if (wheelIsValid && searchIsValid)
                {
                    KeyboardShortcutApplyResult result =
                        KeyboardShortcutService.ApplyToAllInstalledVersions(
                            ModuleData.RevitVersion,
                            ModuleData.RevitVersionName,
                            ModuleData.RibbonTabName,
                            wheelShortcut,
                            searchShortcut,
                            wheelCommandId,
                            searchCommandId);

                    if (result.Success && string.IsNullOrWhiteSpace(result.ErrorMessage))
                    {
                        settings.WheelShortcut = wheelShortcut;
                        settings.CommandSearchShortcut = searchShortcut;
                        settings.AreKeyboardShortcutsConfigured = true;
                        MarkMigrationCompleted(
                            settings,
                            "Migrated",
                            "Старые назначения успешно перенесены.");
                        ClearLegacyHotkeys(settings);
                        UserSettingsService.Save(settings);
                        return false;
                    }

                    settings.KeyboardShortcutMigrationMessage =
                        string.IsNullOrWhiteSpace(result.ErrorMessage)
                            ? "Не удалось записать назначения в KeyboardShortcuts.xml."
                            : result.ErrorMessage;
                }
                else
                {
                    List<string> errors = new List<string>();
                    if (!wheelIsValid && !string.IsNullOrWhiteSpace(wheelError))
                    {
                        errors.Add("Штурвал: " + wheelError);
                    }
                    if (!searchIsValid && !string.IsNullOrWhiteSpace(searchError))
                    {
                        errors.Add("Команды: " + searchError);
                    }

                    settings.KeyboardShortcutMigrationMessage = errors.Count == 0
                        ? "Старые назначения нельзя однозначно перенести."
                        : string.Join(" ", errors.ToArray());
                }

                showNotice = true;
                settings.KeyboardShortcutMigrationStatus = "Failed";
                UserSettingsService.Save(settings);
            }
            catch (Exception ex)
            {
                // Migration must never prevent the command window from opening.
                showNotice = true;
                try
                {
                    settings.LegacyHotkeyMigrationAttempted = true;
                    settings.KeyboardShortcutMigrationStatus = "Failed";
                    settings.KeyboardShortcutMigrationMessage =
                        "Ошибка проверки старых настроек: " + ex.Message;
                    UserSettingsService.Save(settings);
                }
                catch
                {
                    // Ignore an unavailable settings file as well.
                }
            }

            if (showNotice)
            {
                settings.LegacyHotkeyMigrationNoticeShown = true;
                try
                {
                    UserSettingsService.Save(settings);
                }
                catch
                {
                    // The informational dialog is still useful if persistence failed.
                }
            }

            return showNotice;
        }

        private static bool HasLegacyHotkey(LegacyHotkeyGesture gesture)
        {
            return gesture != null
                && ((!string.IsNullOrWhiteSpace(gesture.MouseButton))
                    || (gesture.Keys != null && gesture.Keys.Count != 0));
        }

        private static void ClearLegacyHotkeys(UserSettings settings)
        {
            settings.CommandsWheelHotkey = null;
            settings.CommandSearchHotkey = null;
        }

        private static void MarkMigrationCompleted(
            UserSettings settings,
            string status,
            string message)
        {
            settings.LegacyHotkeyMigrationAttempted = true;
            settings.LegacyHotkeyMigrationNoticeShown = true;
            settings.KeyboardShortcutMigrationStatus = status;
            settings.KeyboardShortcutMigrationMessage = message;
        }
    }
}