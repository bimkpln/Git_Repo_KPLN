using Autodesk.Revit.DB;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;

namespace KPLN_DefaultPanelExtension_Modify.Common
{
    internal sealed class XyzJsonConverter : JsonConverter<XYZ>
    {
        public override XYZ ReadJson(JsonReader reader, Type objectType, XYZ existingValue, bool hasExistingValue, JsonSerializer serializer)
        {
            if (reader.TokenType == JsonToken.Null)
                return null;

            var jo = JObject.Load(reader);

            double x = jo["X"]?.Value<double>() ?? 0;
            double y = jo["Y"]?.Value<double>() ?? 0;
            double z = jo["Z"]?.Value<double>() ?? 0;

            return new XYZ(x, y, z);
        }

        public override void WriteJson(JsonWriter writer, XYZ value, JsonSerializer serializer)
        {
            if (value == null)
            {
                writer.WriteNull();
                return;
            }

            writer.WriteStartObject();
            writer.WritePropertyName("X"); writer.WriteValue(value.X);
            writer.WritePropertyName("Y"); writer.WriteValue(value.Y);
            writer.WritePropertyName("Z"); writer.WriteValue(value.Z);
            writer.WriteEndObject();
        }
    }
}
