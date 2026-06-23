using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace KPLN_CommandsWheel.Services
{
    internal static class JsonSerialization
    {
        internal static T Deserialize<T>(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return default(T);
            }

            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (T)serializer.ReadObject(stream);
            }
        }

        internal static string Serialize(object value)
        {
            if (value == null)
            {
                return "null";
            }

            DataContractJsonSerializer serializer = new DataContractJsonSerializer(value.GetType());
            using (MemoryStream stream = new MemoryStream())
            {
                serializer.WriteObject(stream, value);
                return Format(Encoding.UTF8.GetString(stream.ToArray()));
            }
        }

        private static string Format(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return json;
            }

            StringBuilder builder = new StringBuilder();
            bool inString = false;
            bool escaped = false;
            int indent = 0;

            foreach (char character in json)
            {
                if (escaped)
                {
                    builder.Append(character);
                    escaped = false;
                    continue;
                }

                if (character == '\\' && inString)
                {
                    builder.Append(character);
                    escaped = true;
                    continue;
                }

                if (character == '"')
                {
                    builder.Append(character);
                    inString = !inString;
                    continue;
                }

                if (inString)
                {
                    builder.Append(character);
                    continue;
                }

                switch (character)
                {
                    case '{':
                    case '[':
                        builder.Append(character);
                        AppendLine(builder, ++indent);
                        break;
                    case '}':
                    case ']':
                        AppendLine(builder, --indent);
                        builder.Append(character);
                        break;
                    case ',':
                        builder.Append(character);
                        AppendLine(builder, indent);
                        break;
                    case ':':
                        builder.Append(": ");
                        break;
                    default:
                        if (!char.IsWhiteSpace(character))
                        {
                            builder.Append(character);
                        }
                        break;
                }
            }

            return builder.ToString();
        }

        private static void AppendLine(StringBuilder builder, int indent)
        {
            builder.AppendLine();
            builder.Append(new string(' ', indent * 2));
        }
    }
}