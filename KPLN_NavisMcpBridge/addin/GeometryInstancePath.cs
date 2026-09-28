using System;
using System.Collections.Generic;
using System.Globalization;

namespace KPLN_NavisMcpBridge
{
    internal static class GeometryInstancePath
    {
        internal static string Read(object data)
        {
            var array = data as Array;
            if (array == null || array.Rank != 1 || array.Length == 0)
                throw new InvalidOperationException("Missing or invalid COM instance ArrayData.");
            var parts = new string[array.Length];
            var lower = array.GetLowerBound(0);
            for (int i = 0; i < parts.Length; i++)
            {
                var value = array.GetValue(lower + i);
                if (!(value is byte || value is sbyte || value is short || value is ushort ||
                      value is int || value is uint || value is long || value is ulong))
                    throw new InvalidOperationException("Non-integral COM instance path component.");
                parts[i] = Convert.ToString(value, CultureInfo.InvariantCulture);
            }
            return string.Join("/", parts);
        }

        internal static List<T> Select<T>(IEnumerable<T> candidates, string instanceKey,
            Func<T, object> getPath, out int skipped)
        {
            if (string.IsNullOrEmpty(instanceKey)) throw new ArgumentException("Missing instance key.");
            var selected = new List<T>();
            skipped = 0;
            foreach (var candidate in candidates)
            {
                // Exact full path: a shared tail node, parent prefix or display name is insufficient.
                if (Read(getPath(candidate)) == instanceKey) selected.Add(candidate);
                else skipped++;
            }
            if (selected.Count == 0)
                throw new InvalidOperationException("No fragments match the requested COM instance path.");
            return selected;
        }
    }
}
