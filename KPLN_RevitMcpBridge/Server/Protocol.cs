using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
#if Debug2026 || Revit2026
using System.Text.Json;
#else
using System.Web.Script.Serialization;
#endif

namespace KPLN_RevitMcpBridge.Server
{
    internal sealed class BridgeException : Exception
    {
        public readonly string Code;
        public readonly int Status;
        public BridgeException(string code, string message, int status = 400) : base(message) { Code = code; Status = status; }
    }

    internal static class Json
    {
#if Debug2026 || Revit2026
        private static readonly JsonSerializerOptions SerializerOptions = new JsonSerializerOptions
        {
            IncludeFields = true,
            MaxDepth = 64
        };

        public static string Serialize(object value) => JsonSerializer.Serialize(value, SerializerOptions);

        public static Dictionary<string, object> Parse(string value)
        {
            try
            {
                if (value == null || value.Length > 8 * 1024 * 1024) throw new FormatException();

                using (var document = JsonDocument.Parse(value, new JsonDocumentOptions { MaxDepth = 64 }))
                {
                    if (document.RootElement.ValueKind != JsonValueKind.Object) throw new FormatException();
                    return (Dictionary<string, object>)ReadValue(document.RootElement);
                }
            }
            catch { throw new BridgeException("invalid_json", "Ожидается JSON-объект."); }
        }

        // Сервисы ожидают обычные Dictionary/IList и числа, а не JsonElement.
        private static object ReadValue(JsonElement value)
        {
            switch (value.ValueKind)
            {
                case JsonValueKind.Object:
                    var data = new Dictionary<string, object>();
                    foreach (var property in value.EnumerateObject()) data[property.Name] = ReadValue(property.Value);
                    return data;
                case JsonValueKind.Array:
                    var items = new ArrayList();
                    foreach (var item in value.EnumerateArray()) items.Add(ReadValue(item));
                    return items;
                case JsonValueKind.String:
                    return value.GetString();
                case JsonValueKind.Number:
                    if (value.TryGetInt32(out var intValue)) return intValue;
                    if (value.TryGetInt64(out var longValue)) return longValue;
                    return value.GetDouble();
                case JsonValueKind.True:
                    return true;
                case JsonValueKind.False:
                    return false;
                default:
                    return null;
            }
        }
#else
        private static JavaScriptSerializer Serializer() => new JavaScriptSerializer { MaxJsonLength = 8 * 1024 * 1024, RecursionLimit = 64 };
        public static string Serialize(object value) => Serializer().Serialize(value);
        public static Dictionary<string, object> Parse(string value)
        {
            try { return Serializer().Deserialize<Dictionary<string, object>>(value) ?? throw new Exception(); }
            catch { throw new BridgeException("invalid_json", "Ожидается JSON-объект."); }
        }
#endif

        public static object Get(this IDictionary<string, object> data, string key)
        { object value; return data.TryGetValue(key, out value) ? value : null; }
        public static string Text(this IDictionary<string, object> data, string key, bool required = false)
        {
            var value = data.Get(key);
            if (value != null && !(value is string)) throw new BridgeException("invalid_input", key + ": ожидается строка.");
            var text = value as string;
            if (required && string.IsNullOrEmpty(text)) throw new BridgeException("invalid_input", "Не задано " + key);
            return text;
        }
        public static bool Flag(this IDictionary<string, object> data, string key, bool fallback = false)
        {
            var value = data.Get(key);
            if (value == null) return fallback;
            if (!(value is bool)) throw new BridgeException("invalid_input", key + ": ожидается boolean.");
            return (bool)value;
        }
        public static long Integer(object value, string key)
        {
            long result;
            if (value == null || value is bool || !long.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out result))
                throw new BridgeException("invalid_input", key + ": ожидается целое число.");
            return result;
        }
        public static double Double(object value)
        {
            if (!(value is int || value is long || value is double || value is decimal))
                throw new BridgeException("invalid_value", "Требуется число в единицах Revit.");
            // JavaScriptSerializer reads fractional JSON numbers as Decimal. On
            // .NET Framework, Decimal -> Double can round to an adjacent Double.
            // Parse the invariant decimal text instead, preserving the serialized
            // round-trip value without weakening exact expected_value checks.
            double number = value is decimal
                ? double.Parse(((decimal)value).ToString(CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture)
                : Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (double.IsNaN(number) || double.IsInfinity(number))
                throw new BridgeException("invalid_value", "Требуется конечное число.");
            return number;
        }
        public static int Range(this IDictionary<string, object> data, string key, int fallback, int min, int max)
        {
            var number = data.Get(key) == null ? fallback : Integer(data.Get(key), key);
            if (number < min || number > max) throw new BridgeException("invalid_input", key + ": вне допустимого диапазона.");
            return (int)number;
        }
        public static IList List(this IDictionary<string, object> data, string key, int max = 200)
        {
            var list = data.Get(key) as IList;
            if (list == null || list.Count == 0 || list.Count > max) throw new BridgeException("invalid_input", key + ": ожидается от 1 до " + max + " записей.");
            return list;
        }
        public static object Error(Exception ex)
        {
            var known = ex as BridgeException;
            return new { ok = false, error = new { code = known?.Code ?? "revit_error", message = ex.Message } };
        }
    }
}
