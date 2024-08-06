using Autodesk.Revit.DB;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;

namespace KPLN_Tools.Common.LinkManager
{
    public class CoordinateTypeConverter : JsonConverter<CoordinateType>
    {
        private readonly CoordinateType[] _linkCoordinateTypeColl;

        public CoordinateTypeConverter(CoordinateType[] linkCoordinateTypeColl)
        {
            _linkCoordinateTypeColl = linkCoordinateTypeColl;
        }

        public override CoordinateType ReadJson(JsonReader reader, Type objectType, CoordinateType existingValue, bool hasExistingValue, JsonSerializer serializer)
        {
            var jsonObject = JObject.Load(reader);
            var name = jsonObject["Name"].Value<string>();
            var typeName = jsonObject["Type"].Value<string>();
            var type = _linkCoordinateTypeColl.FirstOrDefault(ct => ct.Name == typeName)?.Type;

            return new CoordinateType(name, type ?? ImportPlacement.Origin);
        }

        public override void WriteJson(JsonWriter writer, CoordinateType value, JsonSerializer serializer)
        {
            var jsonObject = new JObject
            {
                ["Name"] = value.Name,
                ["Type"] = value.Type.ToString() // Or any specific value mapping you have
            };
            jsonObject.WriteTo(writer);
        }
    }
}
