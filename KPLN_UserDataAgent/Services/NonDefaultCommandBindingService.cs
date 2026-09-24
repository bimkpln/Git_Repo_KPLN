using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class NonDefaultCommandBindingService : IDisposable
    {
        private const string CustomCommandIdPrefix = "CustomCtrl_";
        private const int ContinuousScanLimit = 3;
        private static readonly TimeSpan InitialScanInterval = TimeSpan.FromSeconds(5);
        private static readonly TimeSpan PeriodicScanInterval = TimeSpan.FromSeconds(60);
        private static readonly TimeSpan FallbackExecutionWindow = TimeSpan.FromSeconds(60);

        private readonly UIControlledApplication _application;
        private readonly PluginUsageTracker _tracker;
        private readonly ErrorGuard _errorGuard;
        private readonly Dictionary<string, BindingRegistration> _bindings =
            new Dictionary<string, BindingRegistration>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, NonDefaultCommandInfo> _commandsById =
            new Dictionary<string, NonDefaultCommandInfo>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ActiveExecution> _activeExecutions =
            new Dictionary<string, ActiveExecution>(StringComparer.OrdinalIgnoreCase);

        // Index by object identity: unrelated vendors may reuse the same ID.
        private readonly Dictionary<object, NonDefaultCommandInfo> _ribbonItems =
            new Dictionary<object, NonDefaultCommandInfo>(ReferenceComparer.Instance);
        private readonly Dictionary<object, RibbonActivation> _otherActivations =
            new Dictionary<object, RibbonActivation>(ReferenceComparer.Instance);
        private readonly Dictionary<Type, Dictionary<string, PropertyInfo>> _properties =
            new Dictionary<Type, Dictionary<string, PropertyInfo>>();
        private readonly Dictionary<string, bool> _nativeCommandIds =
            new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, bool> _bundledAssemblyCache =
            new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private EventInfo _uiElementActivatedEvent;
        private Delegate _uiElementActivatedHandler;
        private bool _ribbonEventsAttached;

        private Type _componentManagerType;
        private PropertyInfo _ribbonProperty;
        private EventInfo _itemExecutedEvent;
        private Delegate _itemExecutedHandler;
        private int _scanCount;
        private DateTime _lastScanTime = DateTime.MinValue;
        private bool _isDisposed;

        public NonDefaultCommandBindingService(
            UIControlledApplication application,
            PluginUsageTracker tracker,
            ErrorGuard errorGuard)
        {
            _application = application;
            _tracker = tracker;
            _errorGuard = errorGuard;
        }

        public void OnIdling()
        {
            try
            {
                if (_isDisposed)
                    return;

                CloseExecutedScopes();

                if (!ShouldScan())
                    return;

                RegisterMissingBindings();
            }
            catch (Exception exception)
            {
                SafeQueueException("NonDefaultCommandBindingService.Idling", exception);
            }
        }

        public void Dispose()
        {
            try
            {
                _isDisposed = true;

                foreach (BindingRegistration registration in _bindings.Values)
                {
                    try
                    {
                        registration.Detach();
                    }
                    catch
                    {
                    }
                }

                if (_itemExecutedEvent != null && _itemExecutedHandler != null)
                {
                    try
                    {
                        _itemExecutedEvent.RemoveEventHandler(null, _itemExecutedHandler);
                    }
                    catch
                    {
                    }
                }

                if (_uiElementActivatedEvent != null && _uiElementActivatedHandler != null)
                {
                    try { _uiElementActivatedEvent.RemoveEventHandler(null, _uiElementActivatedHandler); }
                    catch { }
                }
                _ribbonItems.Clear();
                _otherActivations.Clear();
                _properties.Clear();
                _nativeCommandIds.Clear();
                _bundledAssemblyCache.Clear();
                _bindings.Clear();
                _commandsById.Clear();

                foreach (ActiveExecution execution in _activeExecutions.Values)
                {
                    try
                    {
                        execution.Scope.Dispose();
                    }
                    catch
                    {
                    }
                }

                _activeExecutions.Clear();
            }
            catch (Exception exception)
            {
                SafeQueueException("NonDefaultCommandBindingService.Dispose", exception);
            }
        }

        private bool ShouldScan()
        {
            DateTime now = DateTime.Now;
            if (_scanCount < ContinuousScanLimit)
            {
                if ((now - _lastScanTime) < InitialScanInterval)
                    return false;
                _scanCount++;
                _lastScanTime = now;
                return true;
            }

            if ((now - _lastScanTime) < PeriodicScanInterval)
                return false;

            _lastScanTime = now;
            return true;
        }

        private void RegisterMissingBindings()
        {
            foreach (NonDefaultCommandInfo commandInfo in CollectNonDefaultCommands())
            {
                if (!commandInfo.IsRegisteredExternal)
                    continue;
                _commandsById[commandInfo.Id] = commandInfo;

                if (_bindings.ContainsKey(commandInfo.Id))
                    continue;

                RevitCommandId commandId = null;
                try
                {
                    commandId = RevitCommandId.LookupCommandId(commandInfo.Id);
                }
                catch
                {
                    commandId = null;
                }

                if (commandId == null || !commandId.CanHaveBinding || commandId.HasBinding)
                    continue;

                AddInCommandBinding binding;
                try
                {
                    binding = _application.CreateAddInCommandBinding(commandId);
                }
                catch
                {
                    continue;
                }

                BindingRegistration registration = new BindingRegistration(
                    commandInfo,
                    binding,
                    (sender, args) => BeginExecution(commandInfo, false),
                    (sender, args) => MarkExecutionCompleted(commandInfo.Id));
                registration.Attach();
                _bindings.Add(commandInfo.Id, registration);
            }
        }

        private IEnumerable<NonDefaultCommandInfo> CollectNonDefaultCommands()
        {
            EnsureRibbonAccessors();
            object ribbon = _ribbonProperty == null ? null : _ribbonProperty.GetValue(null, null);
            Dictionary<string, NonDefaultCommandInfo> result =
                new Dictionary<string, NonDefaultCommandInfo>(StringComparer.OrdinalIgnoreCase);

            _ribbonItems.Clear();
            foreach (object tab in Enumerate(GetPropertyValue(ribbon, "Tabs")))
            {
                string tabName = Clean(FirstString(tab, "Title", "Text", "Name", "Id"));
                foreach (object panel in Enumerate(GetPropertyValue(tab, "Panels")))
                {
                    object panelSource = GetPropertyValue(panel, "Source") ?? panel;
                    string panelName = Clean(FirstString(panelSource, "Title", "Text", "Name", "Id"));
                    object items = GetPropertyValue(panelSource, "Items") ?? GetPropertyValue(panel, "Items");
                    AddItems(items, tabName, panelName, result, 0);
                }
            }

            // Hidden contextual panels are indexed too, so selecting an element
            // does not require another full scan before its first button click.
            List<object> removed = new List<object>();
            foreach (object item in _otherActivations.Keys)
                if (!_ribbonItems.ContainsKey(item)) removed.Add(item);
            foreach (object item in removed) _otherActivations.Remove(item);
            return result.Values;
        }

        private void AddItems(
            object items,
            string tabName,
            string panelName,
            Dictionary<string, NonDefaultCommandInfo> commands,
            int depth)
        {
            if (depth > 8)
                return;

            foreach (object item in Enumerate(items))
            {
                AddCommand(item, tabName, panelName, commands);

                object childItems = GetPropertyValue(item, "Items");
                if (childItems != null)
                    AddItems(childItems, tabName, panelName, commands, depth + 1);

                object source = GetPropertyValue(item, "Source");
                object sourceItems = source == null ? null : GetPropertyValue(source, "Items");
                if (sourceItems != null)
                    AddItems(sourceItems, tabName, panelName, commands, depth + 1);
            }
        }

        private void AddCommand(
            object item,
            string tabName,
            string panelName,
            Dictionary<string, NonDefaultCommandInfo> commands)
        {
            if (item == null || IsMenuContainer(item) || IsBundledAutodeskCommand(item))
                return;

            string id = CleanCommandId(FirstString(item, "Id", "Name"));
            bool registeredExternal = id.StartsWith(CustomCommandIdPrefix, StringComparison.OrdinalIgnoreCase);
            // Preserve existing CustomCtrl_ support, but do not log text boxes,
            // ribbon tabs, row breaks or other UIElementActivated notifications.
            if (!registeredExternal && (!IsExecutableRibbonItem(item) || IsNativeButton(item, id)))
                return;

            string commandName = Clean(FirstString(item, "Text", "ItemText", "Title", "AutomationName", "Name", "Id"));
            if (string.IsNullOrWhiteSpace(commandName))
                commandName = id;

            NonDefaultCommandInfo info = new NonDefaultCommandInfo(id, Clean(tabName), Clean(panelName), commandName);
            _ribbonItems[item] = info;
            if (registeredExternal) commands[id] = info;
        }

        private void BeginExecution(NonDefaultCommandInfo commandInfo, bool isFallbackExecution)
        {
            try
            {
                ActiveExecution previous;
                if (_activeExecutions.TryGetValue(commandInfo.Id, out previous))
                {
                    if ((DateTime.Now - previous.StartedAt).TotalSeconds < 2)
                        return;

                    previous.Scope.Dispose();
                }

                IDisposable scope = _tracker.BeginExecution(
                    commandInfo.TabName,
                    commandInfo.PanelName,
                    commandInfo.CommandName,
                    isFallbackExecution);
                _activeExecutions[commandInfo.Id] = new ActiveExecution(scope, isFallbackExecution);
            }
            catch (Exception exception)
            {
                SafeQueueException("NonDefaultCommandBindingService.BeginExecution", exception);
            }
        }

        private void MarkExecutionCompleted(string commandId)
        {
            try
            {
                ActiveExecution execution;
                if (_activeExecutions.TryGetValue(commandId, out execution))
                    execution.IsCompleted = true;
            }
            catch (Exception exception)
            {
                SafeQueueException("NonDefaultCommandBindingService.Executed", exception);
            }
        }

        private void OnRibbonItemExecuted(object sender, object args)
        {
            HandleRibbonEvent(args, false);
        }

        private void OnRibbonUiElementActivated(object sender, object args)
        {
            HandleRibbonEvent(args, true);
        }

        private void HandleRibbonEvent(object args, bool activation)
        {
            try
            {
                if (_isDisposed) return;
                object item = args == null ? null : GetPropertyValue(args, "Item");
                if (item == null || IsMenuContainer(item) || IsBundledAutodeskCommand(item)) return;
                string id = CleanCommandId(FirstString(item, "Id", "Name"));
                bool registeredExternal = id.StartsWith(CustomCommandIdPrefix, StringComparison.OrdinalIgnoreCase);
                // The existing command binding / ItemExecuted route remains the
                // only owner of registered add-in executions.
                if (activation && registeredExternal) return;
                if (!registeredExternal && (!IsExecutableRibbonItem(item) || IsNativeButton(item, id))) return;

                NonDefaultCommandInfo info;
                if (!_ribbonItems.TryGetValue(item, out info))
                {
                    // Only a previously unseen button causes this refresh.
                    // Cache even unresolved QAT clones below, so repeated clicks
                    // do not rescan the ribbon or lose a newly added panel name.
                    foreach (NonDefaultCommandInfo discovered in CollectNonDefaultCommands())
                        _commandsById[discovered.Id] = discovered;
                    if (!_ribbonItems.TryGetValue(item, out info))
                    {
                        // Event may be from a QAT clone or a just-created button.
                        // Do not guess its source from the currently active tab.
                        if (!registeredExternal || !_commandsById.TryGetValue(id, out info))
                            info = new NonDefaultCommandInfo(id, string.Empty, string.Empty,
                                Clean(FirstString(item, "Text", "ItemText", "Title", "AutomationName", "Name", "Id")));
                        _ribbonItems[item] = info;
                    }
                }
                if (registeredExternal)
                {
                    BeginExecution(info, true);
                    return;
                }

                // UIElementActivated precedes the command and ItemExecuted may
                // follow a long modal dialog. Pair callbacks, not a 2-second
                // debounce that would lose two rapid real clicks.
                RibbonActivation last;
                if (!_otherActivations.TryGetValue(item, out last))
                    _otherActivations[item] = last = new RibbonActivation();
                if (activation)
                {
                    last.AwaitingExecuted = true;
                }
                else if (last.AwaitingExecuted)
                {
                    last.AwaitingExecuted = false;
                    return;
                }
                _tracker.RecordOtherRibbonActivation(info.TabName, info.PanelName, info.CommandName, info.Id);
            }
            catch (Exception exception)
            {
                SafeQueueException("NonDefaultCommandBindingService.RibbonEvent", exception);
            }
        }

        private static bool IsExecutableRibbonItem(object item)
        {
            for (Type type = item.GetType(); type != null; type = type.BaseType)
                if (type.FullName == "Autodesk.Windows.RibbonButton"
                    || type.FullName == "Autodesk.Windows.RibbonMenuItem") return true;
            return false;
        }

        private bool IsMenuContainer(object item)
        {
            // Opening a drop-down is navigation, not a plug-in execution.
            // RibbonMenuButton derives from RibbonButton, so checking only the
            // latter also logged native menus such as Additional Settings.
            for (Type type = item.GetType(); type != null; type = type.BaseType)
            {
                if (type.FullName == "Autodesk.Windows.RibbonListButton")
                    return !GetBoolean(item, "IsSplit", false);
                if (type.FullName == "Autodesk.Windows.RibbonMenuItem")
                    return GetBoolean(item, "HasItems", false);
            }
            return false;
        }

        private bool IsNativeButton(object item, string id)
        {
            if (id.StartsWith(CustomCommandIdPrefix, StringComparison.OrdinalIgnoreCase)) return false;
            object handler = GetPropertyValue(item, "CommandHandler");
            if (handler != null)
            {
                Assembly assembly = handler.GetType().Assembly;
                string name = assembly.GetName().Name;
                // A third-party ICommand handler is positive evidence of an
                // add-in even when the developer has chosen an ID_* identifier.
                if (name != "UIFramework" && name != "RevitAPIUI" && name != "AdWindows")
                    return false;
                // Revit can attach its own generic/dummy handler to an add-in
                // button. The handler's assembly alone does not make it native.
            }
            if (string.IsNullOrEmpty(id)) return false;
            bool native;
            if (_nativeCommandIds.TryGetValue(id, out native)) return native;
            // ID_* alone is not sufficient: also require a registered Revit
            // command. Lookup is cached, including misses, and never binds it.
            native = false;
            if (id.StartsWith("ID_", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    RevitCommandId command = RevitCommandId.LookupCommandId(id);
                    native = command != null && command.Name.StartsWith("ID_", StringComparison.OrdinalIgnoreCase);
                }
                catch { }
            }
            _nativeCommandIds[id] = native;
            return native;
        }

        private bool IsBundledAutodeskCommand(object item)
        {
            // ExternalCommandRibbonButton exposes the actual command assembly.
            // CustomCtrl_ also covers Autodesk's bundled add-ins (Precast,
            // eTransmit, etc.), so a custom ID alone cannot identify a vendor.
            string assemblyPath = FirstString(item, "AssemblyAbsolutePath");
            if (string.IsNullOrWhiteSpace(assemblyPath)) return false;
            bool bundled;
            if (_bundledAssemblyCache.TryGetValue(assemblyPath, out bundled)) return bundled;
            bundled = false;
            try
            {
                string fullPath = Path.GetFullPath(assemblyPath);
                string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                string hostDirectory = Path.GetDirectoryName(typeof(UIControlledApplication).Assembly.Location);
                // Only inspect installed local files. No metadata reads from
                // vendor network shares on Revit's UI thread. Cache per assembly.
                bool installedWithHost = IsInsideDirectory(fullPath, hostDirectory)
                    || (!string.IsNullOrWhiteSpace(programFiles)
                        && IsInsideDirectory(fullPath, Path.Combine(programFiles, "Autodesk")));
                if (installedWithHost && File.Exists(fullPath))
                {
                    string company = FileVersionInfo.GetVersionInfo(fullPath).CompanyName;
                    bundled = !string.IsNullOrWhiteSpace(company)
                        && company.Trim().StartsWith("Autodesk", StringComparison.OrdinalIgnoreCase);
                }
            }
            catch
            {
                // Missing metadata is not proof of a native/bundled command.
            }
            _bundledAssemblyCache[assemblyPath] = bundled;
            return bundled;
        }

        private static bool IsInsideDirectory(string filePath, string directory)
        {
            if (string.IsNullOrWhiteSpace(directory)) return false;
            string prefix = Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            return filePath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }

        private void SafeQueueException(string source, Exception exception)
        {
            try
            {
                _errorGuard?.QueueException(source, exception);
            }
            catch
            {
            }
        }

        private void CloseExecutedScopes()
        {
            DateTime now = DateTime.Now;
            List<string> commandIdsToClose = new List<string>();
            foreach (KeyValuePair<string, ActiveExecution> pair in _activeExecutions)
            {
                if (pair.Value.IsCompleted || pair.Value.ShouldClose(now))
                    commandIdsToClose.Add(pair.Key);
            }

            foreach (string commandId in commandIdsToClose)
            {
                ActiveExecution execution;
                if (!_activeExecutions.TryGetValue(commandId, out execution))
                    continue;

                execution.Scope.Dispose();
                _activeExecutions.Remove(commandId);
            }
        }

        private void EnsureRibbonAccessors()
        {
            if (_ribbonEventsAttached)
                return;

            _componentManagerType = FindType("Autodesk.Windows.ComponentManager");
            if (_componentManagerType == null)
                throw new InvalidOperationException("Autodesk.Windows.ComponentManager not found.");

            _ribbonProperty = _componentManagerType.GetProperty("Ribbon", BindingFlags.Static | BindingFlags.Public);
            if (_ribbonProperty == null)
                throw new InvalidOperationException("Autodesk.Windows.ComponentManager.Ribbon not found.");

            _itemExecutedEvent = _componentManagerType.GetEvent("ItemExecuted", BindingFlags.Static | BindingFlags.Public);
            if (_itemExecutedEvent != null)
            {
                _itemExecutedHandler = CreateEventHandler(_itemExecutedEvent.EventHandlerType, "OnRibbonItemExecuted");
                _itemExecutedEvent.AddEventHandler(null, _itemExecutedHandler);
            }
            _uiElementActivatedEvent = _componentManagerType.GetEvent("UIElementActivated", BindingFlags.Static | BindingFlags.Public);
            if (_uiElementActivatedEvent != null)
            {
                _uiElementActivatedHandler = CreateEventHandler(_uiElementActivatedEvent.EventHandlerType, "OnRibbonUiElementActivated");
                _uiElementActivatedEvent.AddEventHandler(null, _uiElementActivatedHandler);
            }
            _ribbonEventsAttached = true;
        }

        private Delegate CreateEventHandler(Type eventHandlerType, string methodName)
        {
            MethodInfo invokeMethod = eventHandlerType.GetMethod("Invoke");
            ParameterInfo[] parameters = invokeMethod.GetParameters();
            if (parameters.Length != 2)
                throw new InvalidOperationException("Unsupported ItemExecuted event signature.");

            ParameterExpression senderParameter = Expression.Parameter(parameters[0].ParameterType, "sender");
            ParameterExpression argsParameter = Expression.Parameter(parameters[1].ParameterType, "args");
            MethodInfo handlerMethod = GetType().GetMethod(
                methodName,
                BindingFlags.Instance | BindingFlags.NonPublic);

            MethodCallExpression body = Expression.Call(
                Expression.Constant(this),
                handlerMethod,
                Expression.Convert(senderParameter, typeof(object)),
                Expression.Convert(argsParameter, typeof(object)));

            return Expression.Lambda(eventHandlerType, body, senderParameter, argsParameter).Compile();
        }

        private static Type FindType(string typeName)
        {
            foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                Type type = assembly.GetType(typeName, false);
                if (type != null)
                    return type;
            }

            return Type.GetType(typeName + ", AdWindows", false);
        }

        private static IEnumerable<object> Enumerate(object value)
        {
            IEnumerable enumerable = value as IEnumerable;
            if (enumerable == null || value is string)
                yield break;

            foreach (object item in enumerable)
            {
                yield return item;
            }
        }

        private object GetPropertyValue(object value, string propertyName)
        {
            if (value == null || string.IsNullOrWhiteSpace(propertyName))
                return null;

            Type type = value.GetType();
            Dictionary<string, PropertyInfo> accessors;
            if (!_properties.TryGetValue(type, out accessors))
                _properties[type] = accessors = new Dictionary<string, PropertyInfo>(StringComparer.Ordinal);
            PropertyInfo property;
            if (!accessors.TryGetValue(propertyName, out property))
            {
                property = FindPublicProperty(type, propertyName);
                accessors[propertyName] = property;
            }
            if (property == null || property.GetIndexParameters().Length != 0)
                return null;

            try
            {
                return property.GetValue(value, null);
            }
            catch
            {
                return null;
            }
        }

        private static PropertyInfo FindPublicProperty(Type type, string propertyName)
        {
            // UIFramework.TypeSelector hides RibbonList.Items with a different
            // return type. GetProperty(name) throws AmbiguousMatchException for
            // this real Revit control and used to abort the entire ribbon scan.
            // Prefer the most-derived readable non-indexed declaration.
            for (Type current = type; current != null; current = current.BaseType)
            {
                foreach (PropertyInfo property in current.GetProperties(
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly))
                {
                    if (property.Name == propertyName && property.GetGetMethod() != null
                        && property.GetIndexParameters().Length == 0)
                        return property;
                }
            }
            return null;
        }

        private bool GetBoolean(object value, string propertyName, bool fallback)
        {
            object propertyValue = GetPropertyValue(value, propertyName);
            return propertyValue is bool ? (bool)propertyValue : fallback;
        }

        private string FirstString(object value, params string[] propertyNames)
        {
            foreach (string propertyName in propertyNames)
            {
                object propertyValue = GetPropertyValue(value, propertyName);
                string text = propertyValue as string;
                if (!string.IsNullOrWhiteSpace(text))
                    return text;
            }

            return string.Empty;
        }

        private static string Clean(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            string result = value
                .Replace("\r", " ")
                .Replace("\n", " ")
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

        private sealed class NonDefaultCommandInfo
        {
            public NonDefaultCommandInfo(string id, string tabName, string panelName, string commandName)
            {
                Id = id ?? string.Empty;
                TabName = tabName ?? string.Empty;
                PanelName = panelName ?? string.Empty;
                CommandName = commandName ?? string.Empty;
            }

            public bool IsRegisteredExternal => Id.StartsWith(CustomCommandIdPrefix, StringComparison.OrdinalIgnoreCase);
            public string Id { get; private set; }
            public string TabName { get; private set; }
            public string PanelName { get; private set; }
            public string CommandName { get; private set; }
        }

        private sealed class RibbonActivation
        {
            public bool AwaitingExecuted;
        }

        private sealed class ReferenceComparer : IEqualityComparer<object>
        {
            public static readonly ReferenceComparer Instance = new ReferenceComparer();
            public new bool Equals(object x, object y) => ReferenceEquals(x, y);
            public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
        }

        private sealed class BindingRegistration
        {
            private readonly AddInCommandBinding _binding;
            private readonly EventHandler<BeforeExecutedEventArgs> _beforeExecuted;
            private readonly EventHandler<ExecutedEventArgs> _executed;

            public BindingRegistration(
                NonDefaultCommandInfo commandInfo,
                AddInCommandBinding binding,
                EventHandler<BeforeExecutedEventArgs> beforeExecuted,
                EventHandler<ExecutedEventArgs> executed)
            {
                CommandInfo = commandInfo;
                _binding = binding;
                _beforeExecuted = beforeExecuted;
                _executed = executed;
            }

            public NonDefaultCommandInfo CommandInfo { get; private set; }

            public void Attach()
            {
                _binding.BeforeExecuted += _beforeExecuted;
                _binding.Executed += _executed;
            }

            public void Detach()
            {
                _binding.BeforeExecuted -= _beforeExecuted;
                _binding.Executed -= _executed;
            }
        }

        private sealed class ActiveExecution
        {
            public ActiveExecution(IDisposable scope, bool isFallbackExecution)
            {
                Scope = scope;
                StartedAt = DateTime.Now;
                CloseAt = isFallbackExecution
                    ? StartedAt.Add(FallbackExecutionWindow)
                    : (DateTime?)null;
            }

            public IDisposable Scope { get; private set; }
            public bool IsCompleted { get; set; }
            public DateTime StartedAt { get; private set; }
            public DateTime? CloseAt { get; private set; }

            public bool ShouldClose(DateTime now)
            {
                return CloseAt.HasValue && now >= CloseAt.Value;
            }
        }
    }
}
