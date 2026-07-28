using KPLN_CommandsWheel.ExternalCommands;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace KPLN_CommandsWheel.Services
{
    internal sealed class KeyboardShortcutApplyResult
    {
        internal KeyboardShortcutApplyResult()
        {
            AppliedVersions = new List<int>();
            ChangedVersions = new List<int>();
            FilePaths = new List<string>();
        }

        internal bool Success { get; set; }
        internal bool Changed { get; set; }
        internal string FilePath { get; set; }
        internal string BackupPath { get; set; }
        internal string ErrorMessage { get; set; }
        internal List<int> AppliedVersions { get; private set; }
        internal List<int> ChangedVersions { get; private set; }
        internal List<string> FilePaths { get; private set; }
    }

    internal static class KeyboardShortcutService
    {
        private const string PanelName = "Штурвал команд";
        private const string ShortcutSeparator = "#";
        private const string WheelCommandName = "Штурвал";
        private const string SearchCommandName = "Команды";
        internal const string WheelCommandInternalName = "KPLNCommandsWheelRun";
        internal const string SearchCommandInternalName = "KPLNCommandsWheelSearch";
        private static readonly int[] SupportedVersions = { 2020, 2023, 2024, 2025 };

        private sealed class RevitShortcutProfile
        {
            internal int Version { get; set; }
            internal string FilePath { get; set; }
        }

        private static readonly Dictionary<char, char> EnglishToRussian = CreateEnglishToRussianMap();
        private static readonly Dictionary<char, char> RussianToEnglish =
            EnglishToRussian.ToDictionary(pair => pair.Value, pair => pair.Key);

        private static readonly HashSet<string> ReservedShortcuts =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "ALT+F4",
                "CTRL+W",
                "CTRL+C",
                "CTRL+X",
                "CTRL+V",
                "CTRL+Z",
                "CTRL+SHIFT+Z",
                "CTRL+Y",
                "CTRL+N",
                "CTRL+O",
                "CTRL+P",
                "CTRL+S",
                "CTRL+F",
                "CTRL+D",
                "CTRL+F4",
                "SHIFT+W"
            };

        internal static KeyboardShortcutApplyResult Apply(
            int revitVersion,
            string revitVersionName,
            string ribbonTabName,
            string wheelShortcut,
            string searchShortcut,
            string wheelCommandId = null,
            string searchCommandId = null)
        {
            string filePath = GetShortcutFilePath(revitVersion, revitVersionName);
            return ApplyToFile(
                filePath,
                ribbonTabName,
                wheelShortcut,
                searchShortcut,
                wheelCommandId,
                searchCommandId);
        }

        internal static KeyboardShortcutApplyResult ApplyToAllInstalledVersions(
            int currentRevitVersion,
            string currentRevitVersionName,
            string ribbonTabName,
            string wheelShortcut,
            string searchShortcut,
            string currentWheelCommandId = null,
            string currentSearchCommandId = null)
        {
            KeyboardShortcutApplyResult batchResult = new KeyboardShortcutApplyResult();
            List<string> errors = new List<string>();

            foreach (RevitShortcutProfile profile in GetInstalledRevitProfiles(
                currentRevitVersion,
                currentRevitVersionName))
            {
                bool isCurrentVersion = profile.Version == currentRevitVersion;
                KeyboardShortcutApplyResult versionResult = ApplyToFile(
                    profile.FilePath,
                    ribbonTabName,
                    wheelShortcut,
                    searchShortcut,
                    isCurrentVersion ? currentWheelCommandId : null,
                    isCurrentVersion ? currentSearchCommandId : null);

                if (!versionResult.Success)
                {
                    errors.Add(string.Format(
                        "Revit {0}: {1}",
                        profile.Version,
                        versionResult.ErrorMessage));
                    continue;
                }

                if (!batchResult.AppliedVersions.Contains(profile.Version))
                {
                    batchResult.AppliedVersions.Add(profile.Version);
                }

                if (versionResult.Changed
                    && !batchResult.ChangedVersions.Contains(profile.Version))
                {
                    batchResult.ChangedVersions.Add(profile.Version);
                }

                batchResult.Changed |= versionResult.Changed;
                batchResult.FilePaths.Add(profile.FilePath);
            }

            batchResult.AppliedVersions.Sort();
            batchResult.ChangedVersions.Sort();
            batchResult.FilePath = string.Join("; ", batchResult.FilePaths.ToArray());
            batchResult.Success = batchResult.AppliedVersions.Count != 0;
            batchResult.ErrorMessage = errors.Count == 0
                ? null
                : string.Join("\n\n", errors.ToArray());
            return batchResult;
        }

        internal static bool TryReadCurrentShortcuts(
            int revitVersion,
            string revitVersionName,
            string wheelCommandId,
            string searchCommandId,
            out string wheelShortcut,
            out string searchShortcut)
        {
            wheelShortcut = null;
            searchShortcut = null;

            string filePath = GetShortcutFilePath(revitVersion, revitVersionName);
            if (!File.Exists(filePath))
            {
                return false;
            }

            try
            {
                XDocument document = LoadOrCreateDocument(filePath);
                wheelShortcut = ReadConfiguredShortcuts(
                    document,
                    wheelCommandId,
                    WheelCommandInternalName,
                    typeof(CommandsWheel).FullName);
                searchShortcut = ReadConfiguredShortcuts(
                    document,
                    searchCommandId,
                    SearchCommandInternalName,
                    typeof(CommandSearch).FullName);
                return wheelShortcut != null || searchShortcut != null;
            }
            catch
            {
                wheelShortcut = null;
                searchShortcut = null;
                return false;
            }
        }

        private static KeyboardShortcutApplyResult ApplyToFile(
            string filePath,
            string ribbonTabName,
            string wheelShortcut,
            string searchShortcut,
            string wheelCommandId,
            string searchCommandId)
        {
            KeyboardShortcutApplyResult result = new KeyboardShortcutApplyResult
            {
                FilePath = filePath
            };

            try
            {
                string error;
                string wheelVariants = BuildShortcutVariants(wheelShortcut, out error);
                if (error != null)
                {
                    result.ErrorMessage = "Штурвал: " + error;
                    return result;
                }

                string searchVariants = BuildShortcutVariants(searchShortcut, out error);
                if (error != null)
                {
                    result.ErrorMessage = "Команды: " + error;
                    return result;
                }

                string commonShortcut = FindCommonShortcut(
                    wheelVariants,
                    searchVariants);
                if (commonShortcut != null)
                {
                    result.ErrorMessage = string.Format(
                        "Сочетание «{0}» одновременно назначено командам «Штурвал» и «Команды».",
                        commonShortcut);
                    return result;
                }

                string tabName = string.IsNullOrWhiteSpace(ribbonTabName)
                    ? "KPLN"
                    : ribbonTabName.Trim();
                string wheelClassName = typeof(CommandsWheel).FullName;
                string searchClassName = typeof(CommandSearch).FullName;

                XDocument document = LoadOrCreateDocument(filePath);
                XElement root = document.Root;
                if (root == null || !string.Equals(root.Name.LocalName, "Shortcuts", StringComparison.Ordinal))
                {
                    result.ErrorMessage = "Файл KeyboardShortcuts.xml имеет неизвестную структуру.";
                    return result;
                }

                wheelCommandId = ResolveCommandId(
                    root,
                    wheelCommandId,
                    WheelCommandInternalName,
                    tabName);
                searchCommandId = ResolveCommandId(
                    root,
                    searchCommandId,
                    SearchCommandInternalName,
                    tabName);

                string conflict = FindConflict(
                    document,
                    new[] { wheelCommandId, searchCommandId },
                    new[] { wheelClassName, searchClassName },
                    wheelVariants,
                    searchVariants);

                if (!string.IsNullOrWhiteSpace(conflict))
                {
                    result.ErrorMessage = conflict;
                    return result;
                }

                result.Changed = UpdateShortcutItem(
                    root,
                    wheelCommandId,
                    WheelCommandInternalName,
                    wheelClassName,
                    WheelCommandName,
                    tabName,
                    wheelVariants);
                result.Changed |= UpdateShortcutItem(
                    root,
                    searchCommandId,
                    SearchCommandInternalName,
                    searchClassName,
                    SearchCommandName,
                    tabName,
                    searchVariants);
                result.Changed |= RemoveObsoleteGeneratedItems(
                    root,
                    wheelCommandId,
                    wheelClassName,
                    WheelCommandName);
                result.Changed |= RemoveObsoleteGeneratedItems(
                    root,
                    searchCommandId,
                    searchClassName,
                    SearchCommandName);

                if (result.Changed)
                {
                    result.BackupPath = SaveWithBackup(document, filePath);
                }

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = "Не удалось обновить KeyboardShortcuts.xml:\n" + ex.Message;
            }

            return result;
        }

        internal static string FindCommandId(
            IEnumerable<Models.RevitCommandInfo> commands,
            Type commandType)
        {
            if (commands == null || commandType == null || string.IsNullOrWhiteSpace(commandType.FullName))
            {
                return null;
            }

            string suffix = "%" + commandType.FullName;
            string internalName = GetInternalCommandName(commandType);
            Models.RevitCommandInfo command = commands.FirstOrDefault(item =>
                item != null
                && !string.IsNullOrWhiteSpace(item.Id)
                && !string.IsNullOrWhiteSpace(internalName)
                && item.Id.EndsWith(
                    "%" + internalName,
                    StringComparison.OrdinalIgnoreCase));

            if (command == null)
            {
                command = commands.FirstOrDefault(item =>
                    item != null
                    && !string.IsNullOrWhiteSpace(item.Id)
                    && item.Id.EndsWith(suffix, StringComparison.OrdinalIgnoreCase));
            }

            return command == null ? null : command.Id;
        }

        internal static bool TryBuildShortcutPreview(
            string value,
            out string preview,
            out string error)
        {
            string variants = BuildShortcutVariants(value, out error);
            if (error != null)
            {
                preview = null;
                return false;
            }

            preview = string.Join(", ", SplitShortcuts(variants).ToArray());
            return true;
        }

        internal static bool TryNormalizeShortcutInput(
            string value,
            out string englishUpper,
            out string russianUpper,
            out string error)
        {
            error = null;
            string[] configuredShortcuts = SplitConfiguredShortcuts(value).ToArray();
            if (configuredShortcuts.Length == 0)
            {
                englishUpper = string.Empty;
                russianUpper = string.Empty;
                return true;
            }

            if (configuredShortcuts.Length > 1)
            {
                englishUpper = null;
                russianUpper = null;
                error = "можно назначить только одно сочетание или одну последовательность клавиш.";
                return false;
            }

            string russianLower;
            return TryNormalizeSingleShortcutInput(
                configuredShortcuts[0],
                out englishUpper,
                out russianUpper,
                out russianLower,
                out error);
        }

        internal static bool TryNormalizeSingleShortcutInput(
            string value,
            out string englishUpper,
            out string russianUpper,
            out string russianLower,
            out string error)
        {
            string variants = BuildSingleShortcutVariants(value, out error);
            if (error != null)
            {
                englishUpper = null;
                russianUpper = null;
                russianLower = null;
                return false;
            }

            string[] values = SplitShortcuts(variants).ToArray();
            englishUpper = values.FirstOrDefault() ?? string.Empty;
            russianUpper = values.Length > 1 ? values[1] : string.Empty;
            russianLower = values.Length > 2 ? values[2] : russianUpper.ToLowerInvariant();
            return true;
        }

        internal static bool TryConvertLegacyGesture(
            Models.LegacyHotkeyGesture gesture,
            out string shortcut,
            out string error)
        {
            shortcut = string.Empty;
            error = null;

            if (gesture == null
                || ((gesture.Keys == null || gesture.Keys.Count == 0)
                    && string.IsNullOrWhiteSpace(gesture.MouseButton)))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(gesture.MouseButton))
            {
                error = "старое назначение использовало кнопку мыши.";
                return false;
            }

            List<string> modifiers = new List<string>();
            List<string> keys = new List<string>();
            foreach (string rawKey in gesture.Keys ?? new List<string>())
            {
                string modifier = NormalizeLegacyModifier(rawKey);
                if (modifier != null)
                {
                    if (!modifiers.Contains(modifier, StringComparer.OrdinalIgnoreCase))
                    {
                        modifiers.Add(modifier);
                    }
                    continue;
                }

                string key;
                if (!TryNormalizeLegacyAlphaNumericKey(rawKey, out key))
                {
                    error = "старое назначение содержит неподдерживаемую специальную клавишу.";
                    return false;
                }

                if (!keys.Contains(key, StringComparer.OrdinalIgnoreCase))
                {
                    keys.Add(key);
                }
            }

            if (modifiers.Count == 0 || keys.Count != 1)
            {
                error = "старое назначение нельзя однозначно перенести.";
                return false;
            }

            string candidate = BuildModifierPrefix(modifiers) + keys[0];
            string russianUpper;
            string russianLower;
            return TryNormalizeSingleShortcutInput(
                candidate,
                out shortcut,
                out russianUpper,
                out russianLower,
                out error);
        }

        private static string BuildShortcutVariants(string value, out string error)
        {
            error = null;
            string[] configuredShortcuts = SplitConfiguredShortcuts(value).ToArray();
            if (configuredShortcuts.Length == 0)
            {
                return string.Empty;
            }

            if (configuredShortcuts.Length > 1)
            {
                error = "можно назначить только одно сочетание или одну последовательность клавиш.";
                return null;
            }

            return BuildSingleShortcutVariants(configuredShortcuts[0], out error);
        }

        private static string BuildSingleShortcutVariants(string value, out string error)
        {
            error = null;
            string compact = (value ?? string.Empty).Replace(" ", string.Empty).Trim();
            if (compact.Length == 0)
            {
                return string.Empty;
            }

            if (string.Equals(compact, "Shift+Tab", StringComparison.OrdinalIgnoreCase)
                || string.Equals(compact, "Tab", StringComparison.OrdinalIgnoreCase))
            {
                error = "Shift+Tab и Tab зарезервированы самим Revit. Используйте буквенную комбинацию, например KW.";
                return null;
            }

            string[] parts = compact.Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> modifiers = new List<string>();
            string keySequence;

            if (parts.Length == 1)
            {
                keySequence = parts[0];
                if (keySequence.Length < 2 || keySequence.Length > 5)
                {
                    error = "последовательность должна содержать от 2 до 5 букв или цифр.";
                    return null;
                }
            }
            else
            {
                if (parts.Length < 2 || parts.Length > 4)
                {
                    error = "допустимы Ctrl, Shift и Alt плюс одна буква или цифра.";
                    return null;
                }

                keySequence = parts[parts.Length - 1];
                for (int index = 0; index < parts.Length - 1; index++)
                {
                    string modifier = NormalizeModifier(parts[index]);
                    if (modifier == null || modifiers.Contains(modifier, StringComparer.OrdinalIgnoreCase))
                    {
                        error = "допустимы только модификаторы Ctrl, Shift и Alt без повторов.";
                        return null;
                    }

                    modifiers.Add(modifier);
                }

                if (keySequence.Length != 1)
                {
                    error = "после Ctrl, Shift или Alt должна быть ровно одна буква или цифра.";
                    return null;
                }

                if (modifiers.Contains("Alt", StringComparer.OrdinalIgnoreCase)
                    && !modifiers.Contains("Ctrl", StringComparer.OrdinalIgnoreCase)
                    && !modifiers.Contains("Shift", StringComparer.OrdinalIgnoreCase))
                {
                    error = "Alt можно использовать только вместе с Ctrl или Shift.";
                    return null;
                }
            }

            if (!keySequence.All(IsSupportedAlphaNumeric))
            {
                error = "разрешены только английские или русские буквы на буквенных клавишах и цифры.";
                return null;
            }

            string englishLower = ConvertLayout(keySequence, true).ToLowerInvariant();
            string russianLower = ConvertLayout(keySequence, false).ToLowerInvariant();

            if (englishLower.Distinct().Count() != englishLower.Length)
            {
                error = "клавиши внутри последовательности не должны повторяться.";
                return null;
            }

            string modifierPrefix = BuildModifierPrefix(modifiers);
            string reservedCandidate = (modifierPrefix + englishLower.ToUpperInvariant()).TrimEnd('+');
            if (ReservedShortcuts.Contains(reservedCandidate))
            {
                error = string.Format("сочетание {0} зарезервировано самим Revit.", reservedCandidate);
                return null;
            }

            List<string> variants = new List<string>
            {
                modifierPrefix + englishLower.ToUpperInvariant(),
                modifierPrefix + russianLower.ToUpperInvariant(),
                modifierPrefix + russianLower
            };

            return string.Join(
                ShortcutSeparator,
                variants.Where(item => !string.IsNullOrWhiteSpace(item))
                    .Distinct(StringComparer.Ordinal)
                    .ToArray());
        }

        private static string FindConflict(
            XDocument document,
            IEnumerable<string> targetCommandIds,
            IEnumerable<string> targetClassNames,
            params string[] requestedShortcutSets)
        {
            HashSet<string> requested = new HashSet<string>(
                requestedShortcutSets.SelectMany(SplitShortcuts),
                StringComparer.OrdinalIgnoreCase);
            requested.RemoveWhere(string.IsNullOrWhiteSpace);

            if (requested.Count == 0)
            {
                return null;
            }

            HashSet<string> ids = new HashSet<string>(
                targetCommandIds.Where(item => !string.IsNullOrWhiteSpace(item)),
                StringComparer.OrdinalIgnoreCase);
            string[] classNames = targetClassNames
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .ToArray();

            foreach (XElement item in document.Descendants().Where(
                element => string.Equals(element.Name.LocalName, "ShortcutItem", StringComparison.Ordinal)))
            {
                string commandId = GetAttributeValue(item, "CommandId");
                if (ids.Contains(commandId)
                    || classNames.Any(className => CommandIdMatchesClass(commandId, className)))
                {
                    continue;
                }

                string duplicate = SplitShortcuts(GetAttributeValue(item, "Shortcuts"))
                    .FirstOrDefault(requested.Contains);
                if (duplicate == null)
                {
                    continue;
                }

                string commandName = GetAttributeValue(item, "CommandName");
                return string.Format(
                    "Сочетание «{0}» уже назначено команде «{1}». Выберите свободное сочетание.",
                    duplicate,
                    string.IsNullOrWhiteSpace(commandName) ? commandId : commandName);
            }

            return null;
        }

        private static bool UpdateShortcutItem(
            XElement root,
            string commandId,
            string internalCommandName,
            string className,
            string commandName,
            string tabName,
            string shortcuts)
        {
            XElement item = FindShortcutItem(
                root,
                commandId,
                internalCommandName,
                className);

            if (item == null)
            {
                item = new XElement(
                    "ShortcutItem",
                    new XAttribute("CommandName", commandName),
                    new XAttribute("CommandId", commandId),
                    new XAttribute("Shortcuts", shortcuts ?? string.Empty),
                    new XAttribute("Paths", tabName + ">" + PanelName));
                root.Add(item);
                return true;
            }

            bool changed = false;
            changed |= SetAttribute(item, "CommandName", commandName);
            changed |= SetAttribute(item, "CommandId", commandId);
            changed |= SetAttribute(item, "Shortcuts", shortcuts ?? string.Empty);
            changed |= SetAttribute(item, "Paths", tabName + ">" + PanelName);
            return changed;
        }

        private static bool RemoveObsoleteGeneratedItems(
            XElement root,
            string retainedCommandId,
            string className,
            string commandName)
        {
            XElement[] obsoleteItems = root
                .Descendants()
                .Where(element =>
                    string.Equals(element.Name.LocalName, "ShortcutItem", StringComparison.Ordinal)
                    && !string.Equals(
                        GetAttributeValue(element, "CommandId"),
                        retainedCommandId,
                        StringComparison.OrdinalIgnoreCase)
                    && CommandIdMatchesClass(
                        GetAttributeValue(element, "CommandId"),
                        className)
                    && string.Equals(
                        GetAttributeValue(element, "CommandName"),
                        commandName,
                        StringComparison.OrdinalIgnoreCase))
                .ToArray();

            foreach (XElement obsoleteItem in obsoleteItems)
            {
                obsoleteItem.Remove();
            }

            return obsoleteItems.Length != 0;
        }

        private static XDocument LoadOrCreateDocument(string filePath)
        {
            if (File.Exists(filePath))
            {
                if (new FileInfo(filePath).Length == 0
                    || string.IsNullOrWhiteSpace(File.ReadAllText(filePath)))
                {
                    return CreateEmptyShortcutDocument();
                }

                return XDocument.Load(filePath, LoadOptions.PreserveWhitespace);
            }

            return CreateEmptyShortcutDocument();
        }

        private static XDocument CreateEmptyShortcutDocument()
        {
            return new XDocument(
                new XDeclaration("1.0", "utf-8", null),
                new XElement("Shortcuts"));
        }

        private static string ReadConfiguredShortcuts(
            XDocument document,
            string commandId,
            string internalCommandName,
            string className)
        {
            XElement item = FindShortcutItem(
                document.Root,
                commandId,
                internalCommandName,
                className);

            if (item == null)
            {
                return null;
            }

            foreach (string storedShortcut in SplitShortcuts(GetAttributeValue(item, "Shortcuts")))
            {
                string englishUpper;
                string russianUpper;
                string russianLower;
                string error;
                if (!TryNormalizeSingleShortcutInput(
                    storedShortcut,
                    out englishUpper,
                    out russianUpper,
                    out russianLower,
                    out error))
                {
                    continue;
                }

                return englishUpper;
            }

            return string.Empty;
        }

        private static string SaveWithBackup(XDocument document, string filePath)
        {
            string directory = Path.GetDirectoryName(filePath);
            Directory.CreateDirectory(directory);

            string temporaryPath = Path.Combine(directory, "KeyboardShortcuts.kpln.tmp");
            string backupPath = Path.Combine(directory, "KeyboardShortcuts.kpln-backup.xml");

            try
            {
                using (StreamWriter writer = new StreamWriter(
                    temporaryPath,
                    false,
                    new UTF8Encoding(true)))
                {
                    document.Save(writer);
                }

                if (File.Exists(filePath))
                {
                    File.Copy(filePath, backupPath, true);

                    try
                    {
                        File.Replace(temporaryPath, filePath, null, true);
                    }
                    catch
                    {
                        File.Copy(temporaryPath, filePath, true);
                    }
                }
                else
                {
                    File.Move(temporaryPath, filePath);
                    backupPath = null;
                }

                return backupPath;
            }
            finally
            {
                if (File.Exists(temporaryPath))
                {
                    File.Delete(temporaryPath);
                }
            }
        }

        private static string GetShortcutFilePath(int revitVersion, string revitVersionName)
        {
            string revitDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Autodesk",
                "Revit");

            if (!string.IsNullOrWhiteSpace(revitVersionName))
            {
                string exactPath = Path.Combine(revitDirectory, revitVersionName.Trim());
                if (Directory.Exists(exactPath))
                {
                    return Path.Combine(exactPath, "KeyboardShortcuts.xml");
                }
            }

            if (Directory.Exists(revitDirectory))
            {
                Regex versionPattern = new Regex(
                    string.Format(@"(^|\D){0}(\D|$)", revitVersion),
                    RegexOptions.CultureInvariant);

                DirectoryInfo match = new DirectoryInfo(revitDirectory)
                    .EnumerateDirectories()
                    .Where(directory => versionPattern.IsMatch(directory.Name))
                    .OrderByDescending(directory =>
                        directory.Name.StartsWith("Autodesk Revit", StringComparison.OrdinalIgnoreCase))
                    .ThenBy(directory => directory.Name)
                    .FirstOrDefault();

                if (match != null)
                {
                    return Path.Combine(match.FullName, "KeyboardShortcuts.xml");
                }
            }

            string folderName = string.IsNullOrWhiteSpace(revitVersionName)
                ? "Autodesk Revit " + revitVersion
                : revitVersionName.Trim();

            return Path.Combine(revitDirectory, folderName, "KeyboardShortcuts.xml");
        }

        private static List<RevitShortcutProfile> GetInstalledRevitProfiles(
            int currentRevitVersion,
            string currentRevitVersionName)
        {
            List<RevitShortcutProfile> profiles = new List<RevitShortcutProfile>();
            HashSet<string> seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string appDataRevitDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Autodesk",
                "Revit");

            if (Directory.Exists(appDataRevitDirectory))
            {
                foreach (DirectoryInfo directory in new DirectoryInfo(appDataRevitDirectory).EnumerateDirectories())
                {
                    int version;
                    if (TryGetSupportedVersion(directory.Name, out version))
                    {
                        AddProfile(profiles, seenPaths, version, directory.FullName);
                    }
                }
            }

            string programFilesAutodeskDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "Autodesk");

            if (Directory.Exists(programFilesAutodeskDirectory))
            {
                foreach (DirectoryInfo directory in new DirectoryInfo(programFilesAutodeskDirectory).EnumerateDirectories())
                {
                    int version;
                    if (!directory.Name.StartsWith("Revit", StringComparison.OrdinalIgnoreCase)
                        || !TryGetSupportedVersion(directory.Name, out version))
                    {
                        continue;
                    }

                    if (profiles.Any(profile => profile.Version == version))
                    {
                        continue;
                    }

                    string profileDirectory = Path.Combine(
                        appDataRevitDirectory,
                        "Autodesk Revit " + version);
                    AddProfile(profiles, seenPaths, version, profileDirectory);
                }
            }

            string currentFilePath = GetShortcutFilePath(
                currentRevitVersion,
                currentRevitVersionName);
            AddProfile(
                profiles,
                seenPaths,
                currentRevitVersion,
                Path.GetDirectoryName(currentFilePath));

            return profiles
                .OrderBy(profile => profile.Version)
                .ThenBy(profile => profile.FilePath)
                .ToList();
        }

        private static bool TryGetSupportedVersion(string value, out int version)
        {
            foreach (int candidate in SupportedVersions)
            {
                if (Regex.IsMatch(
                    value ?? string.Empty,
                    string.Format(@"(^|\D){0}(\D|$)", candidate),
                    RegexOptions.CultureInvariant))
                {
                    version = candidate;
                    return true;
                }
            }

            version = 0;
            return false;
        }

        private static void AddProfile(
            ICollection<RevitShortcutProfile> profiles,
            ISet<string> seenPaths,
            int version,
            string directoryPath)
        {
            if (!SupportedVersions.Contains(version)
                || string.IsNullOrWhiteSpace(directoryPath))
            {
                return;
            }

            string filePath = Path.Combine(
                Path.GetFullPath(directoryPath),
                "KeyboardShortcuts.xml");
            if (!seenPaths.Add(filePath))
            {
                return;
            }

            profiles.Add(new RevitShortcutProfile
            {
                Version = version,
                FilePath = filePath
            });
        }

        private static string ResolveCommandId(
            XElement root,
            string commandId,
            string internalCommandName,
            string tabName)
        {
            if (!string.IsNullOrWhiteSpace(commandId)
                && CommandIdMatchesInternalName(commandId, internalCommandName))
            {
                return commandId.Trim();
            }

            XElement existingItem = root
                .Descendants()
                .FirstOrDefault(element =>
                    string.Equals(element.Name.LocalName, "ShortcutItem", StringComparison.Ordinal)
                    && CommandIdMatchesInternalName(
                        GetAttributeValue(element, "CommandId"),
                        internalCommandName));

            if (existingItem != null)
            {
                return GetAttributeValue(existingItem, "CommandId");
            }

            return BuildCommandId(tabName, internalCommandName);
        }

        private static XElement FindShortcutItem(
            XElement root,
            string commandId,
            string internalCommandName,
            string className)
        {
            if (root == null)
            {
                return null;
            }

            XElement[] items = root
                .Descendants()
                .Where(element =>
                    string.Equals(element.Name.LocalName, "ShortcutItem", StringComparison.Ordinal))
                .ToArray();

            XElement item = items.FirstOrDefault(element =>
                string.Equals(
                    GetAttributeValue(element, "CommandId"),
                    commandId,
                    StringComparison.OrdinalIgnoreCase));

            if (item != null)
            {
                return item;
            }

            item = items.FirstOrDefault(element =>
                CommandIdMatchesInternalName(
                    GetAttributeValue(element, "CommandId"),
                    internalCommandName));

            return item ?? items.FirstOrDefault(element =>
                CommandIdMatchesClass(
                    GetAttributeValue(element, "CommandId"),
                    className));
        }

        private static string GetInternalCommandName(Type commandType)
        {
            if (commandType == typeof(CommandsWheel))
            {
                return WheelCommandInternalName;
            }

            if (commandType == typeof(CommandSearch))
            {
                return SearchCommandInternalName;
            }

            return null;
        }

        private static string BuildCommandId(string tabName, string internalCommandName)
        {
            return string.Format(
                "CustomCtrl_%CustomCtrl_%{0}%{1}%{2}",
                tabName,
                PanelName,
                internalCommandName);
        }

        private static string NormalizeModifier(string value)
        {
            if (string.Equals(value, "Ctrl", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "Control", StringComparison.OrdinalIgnoreCase))
            {
                return "Ctrl";
            }

            if (string.Equals(value, "Shift", StringComparison.OrdinalIgnoreCase))
            {
                return "Shift";
            }

            if (string.Equals(value, "Alt", StringComparison.OrdinalIgnoreCase))
            {
                return "Alt";
            }

            return null;
        }

        private static string NormalizeLegacyModifier(string value)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (string.Equals(normalized, "Ctrl", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "Control", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "LeftCtrl", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "RightCtrl", StringComparison.OrdinalIgnoreCase))
            {
                return "Ctrl";
            }

            if (string.Equals(normalized, "Shift", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "LeftShift", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "RightShift", StringComparison.OrdinalIgnoreCase))
            {
                return "Shift";
            }

            if (string.Equals(normalized, "Alt", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "Menu", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "LeftAlt", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "RightAlt", StringComparison.OrdinalIgnoreCase))
            {
                return "Alt";
            }

            return null;
        }

        private static bool TryNormalizeLegacyAlphaNumericKey(
            string value,
            out string key)
        {
            key = null;
            string normalized = (value ?? string.Empty).Trim();
            if (normalized.Length == 1 && IsSupportedAlphaNumeric(normalized[0]))
            {
                key = normalized.ToUpperInvariant();
                return true;
            }

            if (normalized.Length == 2
                && (normalized[0] == 'D' || normalized[0] == 'd')
                && char.IsDigit(normalized[1]))
            {
                key = normalized[1].ToString();
                return true;
            }

            return false;
        }

        private static string BuildModifierPrefix(IEnumerable<string> modifiers)
        {
            List<string> ordered = new List<string>();
            if (modifiers.Contains("Ctrl", StringComparer.OrdinalIgnoreCase))
            {
                ordered.Add("Ctrl");
            }
            if (modifiers.Contains("Shift", StringComparer.OrdinalIgnoreCase))
            {
                ordered.Add("Shift");
            }
            if (modifiers.Contains("Alt", StringComparer.OrdinalIgnoreCase))
            {
                ordered.Add("Alt");
            }

            return ordered.Count == 0 ? string.Empty : string.Join("+", ordered) + "+";
        }

        private static bool IsSupportedAlphaNumeric(char value)
        {
            char lower = char.ToLowerInvariant(value);
            return (lower >= '0' && lower <= '9')
                || (lower >= 'a' && lower <= 'z')
                || EnglishToRussian.ContainsValue(lower);
        }

        private static string ConvertLayout(string value, bool toEnglish)
        {
            StringBuilder result = new StringBuilder(value.Length);

            foreach (char sourceCharacter in value)
            {
                char character = char.ToLowerInvariant(sourceCharacter);
                char mapped;

                if (toEnglish && RussianToEnglish.TryGetValue(character, out mapped))
                {
                    result.Append(mapped);
                }
                else if (!toEnglish && EnglishToRussian.TryGetValue(character, out mapped))
                {
                    result.Append(mapped);
                }
                else
                {
                    result.Append(character);
                }
            }

            return result.ToString();
        }

        private static string FindCommonShortcut(string first, string second)
        {
            HashSet<string> values = new HashSet<string>(
                SplitShortcuts(first),
                StringComparer.OrdinalIgnoreCase);
            values.RemoveWhere(string.IsNullOrWhiteSpace);
            return SplitShortcuts(second).FirstOrDefault(values.Contains);
        }

        private static IEnumerable<string> SplitShortcuts(string value)
        {
            return (value ?? string.Empty)
                .Split(new[] { '#' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim())
                .Where(item => item.Length != 0);
        }

        private static IEnumerable<string> SplitConfiguredShortcuts(string value)
        {
            return (value ?? string.Empty)
                .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim())
                .Where(item => item.Length != 0);
        }

        private static bool CommandIdMatchesClass(string commandId, string className)
        {
            return !string.IsNullOrWhiteSpace(commandId)
                && !string.IsNullOrWhiteSpace(className)
                && commandId.EndsWith("%" + className, StringComparison.OrdinalIgnoreCase);
        }

        private static bool CommandIdMatchesInternalName(
            string commandId,
            string internalCommandName)
        {
            return !string.IsNullOrWhiteSpace(commandId)
                && !string.IsNullOrWhiteSpace(internalCommandName)
                && commandId.EndsWith(
                    "%" + internalCommandName,
                    StringComparison.OrdinalIgnoreCase);
        }

        private static string GetAttributeValue(XElement element, string attributeName)
        {
            XAttribute attribute = element.Attribute(attributeName);
            return attribute == null ? string.Empty : attribute.Value;
        }

        private static bool SetAttribute(XElement element, string name, string value)
        {
            XAttribute attribute = element.Attribute(name);
            if (attribute != null && string.Equals(attribute.Value, value, StringComparison.Ordinal))
            {
                return false;
            }

            element.SetAttributeValue(name, value);
            return true;
        }

        private static Dictionary<char, char> CreateEnglishToRussianMap()
        {
            const string english = "qwertyuiopasdfghjklzxcvbnm";
            const string russian = "йцукенгшщзфывапролдячсмить";
            Dictionary<char, char> result = new Dictionary<char, char>();

            for (int index = 0; index < english.Length; index++)
            {
                result[english[index]] = russian[index];
            }

            return result;
        }
    }
}