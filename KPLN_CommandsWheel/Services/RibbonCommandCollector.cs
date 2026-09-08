using Autodesk.Revit.UI;
using Autodesk.Windows;
using KPLN_CommandsWheel.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_CommandsWheel.Services
{
    internal static class RibbonCommandCollector
    {
        private static Dictionary<string, RevitCommandInfo> _cachedCommands;
        private static string _catalogPath;
        private static bool _catalogDirty;

        internal static List<RevitCommandInfo> Collect(UIApplication uiapp)
        {
            EnsureCatalogLoaded(uiapp);

            // Revit can materialize contextual controls after the first scan.
            // Scan their IDs again, but resolve and store each command only once.
            foreach (RibbonCommandItem item in EnumerateRibbonItems())
            {
                AddCommand(item);
            }

            SaveCatalog();
            return _cachedCommands.Values
                .OrderBy(command => command.Name)
                .ThenBy(command => command.TabName)
                .ThenBy(command => command.PanelName)
                .ToList();
        }

        internal static bool IsAvailableInCurrentContext(string commandId)
        {
            object item;
            ICommand handler;
            return TryFindAvailableCommand(commandId, false, out item, out handler);
        }

        internal static RevitCommandId ResolveCommandId(string commandId)
        {
            RibbonCommandItem item = EnumerateRibbonItems()
                .FirstOrDefault(candidate => MatchesCommand(candidate.Item, commandId));
            string nativeId = item == null ? commandId : GetNativeCommandId(item.Item);
            try
            {
                return RevitCommandId.LookupCommandId(nativeId);
            }
            catch
            {
                return null;
            }
        }

        internal static bool TryExecuteRibbonCommand(string commandId)
        {
            // Called on Revit's UI dispatcher after the ExternalEvent has returned.
            // Find a live button and use the same parameter and availability check
            // as Revit's ribbon. Never retain a handler from an earlier context.
            object item;
            ICommand handler;
            if (!TryFindAvailableCommand(commandId, true, out item, out handler))
            {
                return false;
            }

            try
            {
                handler.Execute(item);
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("KPLN Commands Wheel: ribbon command failed. " + ex.Message);
                return false;
            }
        }

        private static bool TryFindAvailableCommand(
            string commandId, bool requireHandler, out object item, out ICommand handler)
        {
            item = null;
            handler = null;
            foreach (RibbonCommandItem candidate in EnumerateRibbonItems())
            {
                if (!(candidate.Item is Autodesk.Windows.RibbonCommandItem)
                    || !candidate.IsAvailable || !MatchesCommand(candidate.Item, commandId))
                {
                    continue;
                }

                ICommand candidateHandler = GetCommandHandler(candidate.Item);
                if (candidateHandler == null && requireHandler)
                {
                    continue;
                }

                try
                {
                    if (candidateHandler != null && !candidateHandler.CanExecute(candidate.Item))
                    {
                        continue;
                    }
                }
                catch
                {
                    continue;
                }

                item = candidate.Item;
                handler = candidateHandler;
                return true;
            }

            return false;
        }

        private static ICommand GetCommandHandler(object item)
        {
            ICommand handler = GetPropertyValue(item, "CommandHandler") as ICommand;
            if (handler != null)
            {
                return handler;
            }

            // Native buttons normally use Revit's shared ribbon handler.
            // Resolve only public members of an already loaded UIFramework assembly.
            foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (!string.Equals(assembly.GetName().Name, "UIFramework", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    Type type = assembly.GetType("UIFramework.RibbonGlobalHandler", false);
                    PropertyInfo property = type == null ? null
                        : type.GetProperty("Command", BindingFlags.Public | BindingFlags.Static);
                    return property == null ? null : property.GetValue(null, null) as ICommand;
                }
                catch
                {
                    return null;
                }
            }

            return null;
        }

        private static bool MatchesCommand(object item, string commandId)
        {
            return string.Equals(GetNativeCommandId(item), commandId, StringComparison.OrdinalIgnoreCase)
                // Preserve favorites and wheel slots saved with an old control ID.
                || string.Equals(GetControlId(item), commandId, StringComparison.OrdinalIgnoreCase);
        }

        private static string GetControlId(object item)
        {
            return CleanCommandId(FirstString(item, "Id", "Name"));
        }

        private static string GetNativeCommandId(object item)
        {
            // UIFramework.ControlHelper stores the command separately in Tag.CommandId.
            // The visual Id may describe a ribbon control and need not be postable.
            string commandId = FirstString(GetPropertyValue(item, "Tag"), "CommandId");
            if (string.IsNullOrWhiteSpace(commandId))
            {
                commandId = FirstString(item, "CommandId");
            }

            return string.IsNullOrWhiteSpace(commandId) ? GetControlId(item) : CleanCommandId(commandId);
        }

        internal static void ClearCache()
        {
            SaveCatalog();
            _cachedCommands = null;
            _catalogPath = null;
            _catalogDirty = false;
        }

        private static void EnsureCatalogLoaded(UIApplication uiapp)
        {
            if (_cachedCommands != null)
            {
                return;
            }

            _cachedCommands = new Dictionary<string, RevitCommandInfo>(StringComparer.OrdinalIgnoreCase);
            _catalogPath = Path.Combine(
                UserSettingsService.SettingsDirectory,
                string.Format("commands.{0}.{1}.json",
                    uiapp.Application.VersionNumber,
                    uiapp.Application.Language));

            try
            {
                if (!File.Exists(_catalogPath))
                {
                    return;
                }

                List<RevitCommandInfo> saved = JsonSerialization.Deserialize<List<RevitCommandInfo>>(
                    File.ReadAllText(_catalogPath, Encoding.UTF8));
                foreach (RevitCommandInfo command in saved ?? new List<RevitCommandInfo>())
                {
                    if (command == null || string.IsNullOrWhiteSpace(command.Id)
                        || string.IsNullOrWhiteSpace(command.Name))
                    {
                        continue;
                    }

                    command.Id = CleanCommandId(command.Id);
                    if (!_cachedCommands.ContainsKey(command.Id))
                    {
                        RestoreIcon(command);
                        _cachedCommands.Add(command.Id, command);
                    }
                }
            }
            catch (Exception ex)
            {
                // An unreadable catalog must not prevent command discovery.
                Debug.WriteLine("KPLN Commands Wheel: cannot load command catalog. " + ex.Message);
            }
        }

        private static void SaveCatalog()
        {
            if (!_catalogDirty || _cachedCommands == null || string.IsNullOrWhiteSpace(_catalogPath))
            {
                return;
            }

            string temporaryPath = _catalogPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                Directory.CreateDirectory(UserSettingsService.SettingsDirectory);
                File.WriteAllText(temporaryPath, JsonSerialization.Serialize(
                    _cachedCommands.Values.OrderBy(command => command.Id).ToList()), Encoding.UTF8);

                if (File.Exists(_catalogPath))
                {
                    File.Replace(temporaryPath, _catalogPath, null);
                }
                else
                {
                    File.Move(temporaryPath, _catalogPath);
                }

                _catalogDirty = false;
            }
            catch (Exception ex)
            {
                // Keep the in-memory catalog and retry on the next collection.
                Debug.WriteLine("KPLN Commands Wheel: cannot save command catalog. " + ex.Message);
            }
            finally
            {
                try
                {
                    if (File.Exists(temporaryPath))
                    {
                        File.Delete(temporaryPath);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("KPLN Commands Wheel: cannot remove catalog temporary file. " + ex.Message);
                }
            }
        }

        private static IEnumerable<RibbonCommandItem> EnumerateRibbonItems()
        {
            object ribbon = ComponentManager.Ribbon;
            foreach (object tab in Enumerate(GetPropertyValue(ribbon, "Tabs")))
            {
                bool tabAvailable = IsAvailable(tab);
                string tabName = FirstString(tab, "Title", "Text", "Name", "Id");
                foreach (object panel in Enumerate(GetPropertyValue(tab, "Panels")))
                {
                    object panelSource = GetPropertyValue(panel, "Source") ?? panel;
                    string panelName = FirstString(panelSource, "Title", "Text", "Name", "Id");
                    object items = GetPanelItems(panelSource);
                    foreach (RibbonCommandItem item in EnumerateItems(
                        items, tabName, panelName, tabAvailable && GetBoolean(panel, "IsEnabled", true), 0))
                    {
                        yield return item;
                    }
                }
            }
        }

        private static IEnumerable<object> GetPanelItems(object source)
        {
            // Slide-out and dynamically generated items are separate from Items.
            HashSet<object> seen = new HashSet<object>(ReferenceEqualityComparer.Instance);
            foreach (string property in new[] { "Items", "ItemsView", "SlideOutPanelItemsView" })
            {
                foreach (object item in Enumerate(GetPropertyValue(source, property)))
                {
                    if (item != null && seen.Add(item))
                    {
                        yield return item;
                    }
                }
            }

            object launcher = GetPropertyValue(source, "DialogLauncher");
            if (launcher != null && seen.Add(launcher))
            {
                yield return launcher;
            }
        }

        private static IEnumerable<RibbonCommandItem> EnumerateItems(
            object items,
            string tabName,
            string panelName,
            bool parentAvailable,
            int depth)
        {
            if (depth > 8)
            {
                yield break;
            }

            foreach (object item in Enumerate(items))
            {
                // RibbonFoldPanel is an invisible layout container by design.
                // Its visibility must not disable the actual buttons inside it.
                bool available = parentAvailable && (item is RibbonRowPanel
                    ? GetBoolean(item, "IsEnabled", true) : IsAvailable(item));
                yield return new RibbonCommandItem
                {
                    Item = item,
                    TabName = tabName,
                    PanelName = panelName,
                    IsAvailable = available
                };

                object childItems = GetPropertyValue(item, "Items");
                if (childItems != null)
                {
                    foreach (RibbonCommandItem child in EnumerateItems(
                        childItems, tabName, panelName, available, depth + 1))
                    {
                        yield return child;
                    }
                }

                object source = GetPropertyValue(item, "Source");
                object sourceItems = source == null ? null : GetPropertyValue(source, "Items");
                if (sourceItems != null && !ReferenceEquals(sourceItems, childItems))
                {
                    foreach (RibbonCommandItem child in EnumerateItems(
                        sourceItems, tabName, panelName, available, depth + 1))
                    {
                        yield return child;
                    }
                }
            }
        }

        private static void AddCommand(RibbonCommandItem ribbonItem)
        {
            object item = ribbonItem.Item;
            if (!(item is Autodesk.Windows.RibbonCommandItem))
            {
                return;
            }

            string id = GetNativeCommandId(item);
            if (string.IsNullOrWhiteSpace(id))
            {
                return;
            }

            RevitCommandInfo existing;
            if (_cachedCommands.TryGetValue(id, out existing)
                || _cachedCommands.TryGetValue(GetControlId(item), out existing))
            {
                CaptureIcon(existing, item);
                return;
            }

            string name = Clean(FirstString(item, "Text", "ItemText", "Title", "Name"));
            if (string.IsNullOrWhiteSpace(name) || string.Equals(id, name, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            // Keep real ribbon commands even when LookupCommandId cannot resolve
            // them. Native contextual commands may require the ribbon handler.
            RevitCommandInfo command = new RevitCommandInfo
            {
                Id = id,
                Name = name,
                TabName = Clean(ribbonItem.TabName),
                PanelName = Clean(ribbonItem.PanelName),
                Tooltip = Clean(FirstString(item, "Description", "ToolTip", "HelpText"))
            };
            CaptureIcon(command, item);
            _cachedCommands.Add(id, command);
            _catalogDirty = true;
        }

        private static bool IsAvailable(object item)
        {
            return item != null
                && GetBoolean(item, "IsVisible", true)
                && GetBoolean(item, "IsEnabled", true);
        }

        private sealed class RibbonCommandItem
        {
            internal object Item { get; set; }
            internal string TabName { get; set; }
            internal string PanelName { get; set; }
            internal bool IsAvailable { get; set; }
        }

        private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
        {
            internal static readonly ReferenceEqualityComparer Instance = new ReferenceEqualityComparer();
            public new bool Equals(object left, object right) { return ReferenceEquals(left, right); }
            public int GetHashCode(object value)
            {
                return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
            }
        }

        private static void RestoreIcon(RevitCommandInfo command)
        {
            if (string.IsNullOrWhiteSpace(command.IconPngBase64))
            {
                return;
            }

            try
            {
                using (MemoryStream stream = new MemoryStream(Convert.FromBase64String(command.IconPngBase64)))
                {
                    PngBitmapDecoder decoder = new PngBitmapDecoder(
                        stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                    BitmapSource image = decoder.Frames[0];
                    if (image.CanFreeze)
                    {
                        image.Freeze();
                    }

                    // OnLoad keeps pixels usable after closing the stream, even
                    // before Revit creates the corresponding contextual button.
                    command.RibbonImage = image;
                }
            }
            catch (Exception ex)
            {
                // One damaged icon must not discard the command or the catalog.
                // Capture it again when its live button is encountered.
                command.IconPngBase64 = null;
                command.RibbonImage = null;
                _catalogDirty = true;
                Debug.WriteLine("KPLN Commands Wheel: cannot restore icon. " + ex.Message);
            }
        }

        private static void CaptureIcon(RevitCommandInfo command, object item)
        {
            if (!string.IsNullOrWhiteSpace(command.IconPngBase64))
            {
                return;
            }

            // Old catalogs receive icons on the next encounter. Missing or not
            // yet initialized images are retried; only a successful capture sticks.
            foreach (string property in new[] { "LargeImage", "Image" })
            {
                ImageSource source = GetPropertyValue(item, property) as ImageSource;
                if (source == null)
                {
                    continue;
                }

                if (command.RibbonImage == null)
                {
                    command.RibbonImage = source;
                }

                string encoded = EncodeIcon(source);
                if (encoded != null)
                {
                    command.IconPngBase64 = encoded;
                    command.RibbonImage = source;
                    _catalogDirty = true;
                    return;
                }
            }
        }

        private static string EncodeIcon(ImageSource source)
        {
            try
            {
                BitmapSource bitmap = source as BitmapSource;
                if (bitmap != null && bitmap.IsDownloading)
                {
                    return null;
                }

                // Preserve small native bitmaps. Rasterize vector icons and cap
                // oversized images so the persistent catalog stays compact.
                if (bitmap == null || bitmap.PixelWidth > 128 || bitmap.PixelHeight > 128)
                {
                    double width = bitmap == null ? source.Width : bitmap.PixelWidth;
                    double height = bitmap == null ? source.Height : bitmap.PixelHeight;
                    if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0
                        || double.IsNaN(height) || double.IsInfinity(height) || height <= 0)
                    {
                        return null;
                    }

                    double scale = (bitmap == null ? 64.0 : 128.0) / Math.Max(width, height);
                    int pixelWidth = Math.Max(1, (int)Math.Round(width * scale));
                    int pixelHeight = Math.Max(1, (int)Math.Round(height * scale));
                    DrawingVisual visual = new DrawingVisual();
                    using (DrawingContext context = visual.RenderOpen())
                    {
                        context.DrawImage(source, new System.Windows.Rect(0, 0, pixelWidth, pixelHeight));
                    }

                    RenderTargetBitmap rendered = new RenderTargetBitmap(
                        pixelWidth, pixelHeight, 96, 96, PixelFormats.Pbgra32);
                    rendered.Render(visual);
                    bitmap = rendered;
                }

                if (bitmap.PixelWidth <= 0 || bitmap.PixelHeight <= 0)
                {
                    return null;
                }

                PngBitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmap));
                using (MemoryStream stream = new MemoryStream())
                {
                    encoder.Save(stream);
                    return Convert.ToBase64String(stream.ToArray());
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("KPLN Commands Wheel: cannot capture icon. " + ex.Message);
                return null;
            }
        }

        private static IEnumerable<object> Enumerate(object value)
        {
            IEnumerable enumerable = value as IEnumerable;
            if (enumerable == null || value is string)
            {
                yield break;
            }

            foreach (object item in enumerable)
            {
                yield return item;
            }
        }

        private static object GetPropertyValue(object value, string propertyName)
        {
            if (value == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return null;
            }

            try
            {
                PropertyInfo property = value.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
                if (property == null || property.GetIndexParameters().Length != 0)
                {
                    return null;
                }

                return property.GetValue(value, null);
            }
            catch
            {
                return null;
            }
        }

        private static bool GetBoolean(object value, string propertyName, bool fallback)
        {
            object propertyValue = GetPropertyValue(value, propertyName);
            if (propertyValue is bool)
            {
                return (bool)propertyValue;
            }

            return fallback;
        }

        private static string FirstString(object value, params string[] propertyNames)
        {
            foreach (string propertyName in propertyNames)
            {
                object propertyValue = GetPropertyValue(value, propertyName);
                string text = propertyValue as string;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    return text;
                }
            }

            return string.Empty;
        }

        private static string Clean(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string result = value
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("_", string.Empty)
                .Trim();

            while (result.Contains("  "))
            {
                result = result.Replace("  ", " ");
            }

            return result;
        }

        private static string CleanCommandId(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }
    }
}