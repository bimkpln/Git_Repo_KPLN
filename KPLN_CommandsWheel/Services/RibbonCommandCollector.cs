using Autodesk.Revit.UI;
using Autodesk.Windows;
using KPLN_CommandsWheel.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Media;

namespace KPLN_CommandsWheel.Services
{
    internal static class RibbonCommandCollector
    {
        internal static List<RevitCommandInfo> Collect(UIApplication uiapp)
        {
            List<RevitCommandInfo> commands = new List<RevitCommandInfo>();
            HashSet<string> seenIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            object ribbon = ComponentManager.Ribbon;
            foreach (object tab in Enumerate(GetPropertyValue(ribbon, "Tabs")))
            {
                if (!GetBoolean(tab, "IsVisible", true))
                {
                    continue;
                }

                string tabName = FirstString(tab, "Title", "Text", "Name", "Id");
                foreach (object panel in Enumerate(GetPropertyValue(tab, "Panels")))
                {
                    if (!GetBoolean(panel, "IsVisible", true))
                    {
                        continue;
                    }

                    object panelSource = GetPropertyValue(panel, "Source") ?? panel;
                    string panelName = FirstString(panelSource, "Title", "Text", "Name", "Id");
                    object items = GetPropertyValue(panelSource, "Items") ?? GetPropertyValue(panel, "Items");
                    AddItems(items, tabName, panelName, uiapp, commands, seenIds, 0);
                }
            }

            return commands
                .OrderBy(command => command.Name)
                .ThenBy(command => command.TabName)
                .ThenBy(command => command.PanelName)
                .ToList();
        }

        private static void AddItems(
            object items,
            string tabName,
            string panelName,
            UIApplication uiapp,
            List<RevitCommandInfo> commands,
            HashSet<string> seenIds,
            int depth)
        {
            if (depth > 8)
            {
                return;
            }

            foreach (object item in Enumerate(items))
            {
                AddCommand(item, tabName, panelName, uiapp, commands, seenIds);

                object childItems = GetPropertyValue(item, "Items");
                if (childItems != null)
                {
                    AddItems(childItems, tabName, panelName, uiapp, commands, seenIds, depth + 1);
                }

                object source = GetPropertyValue(item, "Source");
                object sourceItems = source == null ? null : GetPropertyValue(source, "Items");
                if (sourceItems != null)
                {
                    AddItems(sourceItems, tabName, panelName, uiapp, commands, seenIds, depth + 1);
                }
            }
        }

        private static void AddCommand(
            object item,
            string tabName,
            string panelName,
            UIApplication uiapp,
            List<RevitCommandInfo> commands,
            HashSet<string> seenIds)
        {
            if (item == null || !GetBoolean(item, "IsVisible", true))
            {
                return;
            }

            string id = CleanCommandId(FirstString(item, "Id", "Name"));
            string name = Clean(FirstString(item, "Text", "ItemText", "Title", "Name"));
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(name) || string.Equals(id, name, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            RevitCommandId commandId = null;
            try
            {
                commandId = RevitCommandId.LookupCommandId(id);
            }
            catch
            {
                commandId = null;
            }

            if (commandId == null || !seenIds.Add(id))
            {
                return;
            }

            bool canPost = false;
            try
            {
                canPost = uiapp != null && uiapp.CanPostCommand(commandId);
            }
            catch
            {
                canPost = false;
            }

            commands.Add(new RevitCommandInfo
            {
                Id = id,
                Name = name,
                TabName = Clean(tabName),
                PanelName = Clean(panelName),
                Tooltip = Clean(FirstString(item, "Description", "ToolTip", "HelpText")),
                CanPost = canPost,
                RibbonImage = GetImageSource(item)
            });
        }

        private static ImageSource GetImageSource(object item)
        {
            return GetPropertyValue(item, "LargeImage") as ImageSource
                ?? GetPropertyValue(item, "Image") as ImageSource;
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

            PropertyInfo property = value.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
            if (property == null || property.GetIndexParameters().Length != 0)
            {
                return null;
            }

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