using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class NonDefaultCommandBindingService : IDisposable
    {
        private const string CustomCommandIdPrefix = "CustomCtrl_";
        private const int ContinuousScanLimit = 20;
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

            foreach (object tab in Enumerate(GetPropertyValue(ribbon, "Tabs")))
            {
                if (!GetBoolean(tab, "IsVisible", true))
                    continue;

                string tabName = Clean(FirstString(tab, "Title", "Text", "Name", "Id"));
                foreach (object panel in Enumerate(GetPropertyValue(tab, "Panels")))
                {
                    if (!GetBoolean(panel, "IsVisible", true))
                        continue;

                    object panelSource = GetPropertyValue(panel, "Source") ?? panel;
                    string panelName = Clean(FirstString(panelSource, "Title", "Text", "Name", "Id"));
                    object items = GetPropertyValue(panelSource, "Items") ?? GetPropertyValue(panel, "Items");
                    AddItems(items, tabName, panelName, result, 0);
                }
            }

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

        private static void AddCommand(
            object item,
            string tabName,
            string panelName,
            Dictionary<string, NonDefaultCommandInfo> commands)
        {
            if (item == null || !GetBoolean(item, "IsVisible", true))
                return;

            string id = CleanCommandId(FirstString(item, "Id", "Name"));
            if (!id.StartsWith(CustomCommandIdPrefix, StringComparison.OrdinalIgnoreCase))
                return;

            string commandName = Clean(FirstString(item, "Text", "ItemText", "Title", "AutomationName", "Name", "Id"));
            if (string.IsNullOrWhiteSpace(commandName))
                commandName = id;

            commands[id] = new NonDefaultCommandInfo(id, Clean(tabName), Clean(panelName), commandName);
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
            try
            {
                object item = args == null ? null : GetPropertyValue(args, "Item");
                if (item == null)
                    return;

                string commandId = CleanCommandId(FirstString(item, "Id", "Name"));
                if (string.IsNullOrWhiteSpace(commandId))
                    return;

                NonDefaultCommandInfo commandInfo;
                if (!_commandsById.TryGetValue(commandId, out commandInfo))
                {
                    foreach (NonDefaultCommandInfo refreshedCommandInfo in CollectNonDefaultCommands())
                    {
                        _commandsById[refreshedCommandInfo.Id] = refreshedCommandInfo;
                    }

                    if (!_commandsById.TryGetValue(commandId, out commandInfo))
                        return;
                }

                BeginExecution(commandInfo, true);
            }
            catch (Exception exception)
            {
                SafeQueueException("NonDefaultCommandBindingService.ItemExecuted", exception);
            }
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
            if (_componentManagerType != null)
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
                _itemExecutedHandler = CreateEventHandler(_itemExecutedEvent.EventHandlerType);
                _itemExecutedEvent.AddEventHandler(null, _itemExecutedHandler);
            }
        }

        private Delegate CreateEventHandler(Type eventHandlerType)
        {
            MethodInfo invokeMethod = eventHandlerType.GetMethod("Invoke");
            ParameterInfo[] parameters = invokeMethod.GetParameters();
            if (parameters.Length != 2)
                throw new InvalidOperationException("Unsupported ItemExecuted event signature.");

            ParameterExpression senderParameter = Expression.Parameter(parameters[0].ParameterType, "sender");
            ParameterExpression argsParameter = Expression.Parameter(parameters[1].ParameterType, "args");
            MethodInfo handlerMethod = GetType().GetMethod(
                "OnRibbonItemExecuted",
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

        private static object GetPropertyValue(object value, string propertyName)
        {
            if (value == null || string.IsNullOrWhiteSpace(propertyName))
                return null;

            PropertyInfo property = value.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
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

        private static bool GetBoolean(object value, string propertyName, bool fallback)
        {
            object propertyValue = GetPropertyValue(value, propertyName);
            return propertyValue is bool ? (bool)propertyValue : fallback;
        }

        private static string FirstString(object value, params string[] propertyNames)
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

            public string Id { get; private set; }
            public string TabName { get; private set; }
            public string PanelName { get; private set; }
            public string CommandName { get; private set; }
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