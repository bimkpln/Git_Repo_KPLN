using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using F = System.Windows.Forms;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    [DataContract]
    internal sealed class FinishRules
    {
        [DataMember]
        public string Source = "ASML_Принадлежность к шкафу", Target = "ASML_Коэффициент отделки", Finish = "ASML_Тип отделки";
        [DataMember]
        public string Marker = "ЩК", SectionMarker = "С", AllowedSections = "", ExcludeFamily = "";
        [DataMember]
        public string BaseTemplate = "Тип_{type}", SumTemplate = "SUM_{type}";
        [DataMember]
        public bool UseSections = false, UseSum = true, Fallback = true, SectionInText = true, Remember = true;


        internal List<string> Sections()
        {
            return Regex.Split(AllowedSections ?? "", "[,;]+").Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
        }


        internal string GlobalName(string template, string section, string type)
        {
            return template.Replace("{section}", section).Replace("{type}", type);
        }


        internal Regex Pattern()
        {
            string sectionPattern = string.Equals(SectionMarker, "С", StringComparison.OrdinalIgnoreCase) ? "[СсCc]" : Regex.Escape(SectionMarker);
            return new Regex((UseSections ? sectionPattern + @"\s*(?<section>[0-9]+)\s*" : "") + Regex.Escape(Marker) + @"\s*(?<type>[0-9]+(?:\.[0-9]+)*)(?![0-9.])", RegexOptions.IgnoreCase);
        }


        internal List<string> FormulaTypes(string formula, string section)
        {
            if (string.IsNullOrWhiteSpace(formula))
                throw new InvalidOperationException("У суммарного параметра нет формулы");
            // Разрешены только ссылки на обычные типы, соединённые знаком +, и группирующие скобки.
            // Не интерпретируем множители, условия или вложенные SUM как состав типов.
            string expression = formula.Trim().TrimStart('=').Trim();
            string token = Regex.Escape(BaseTemplate.Replace("{section}", section)).Replace(Regex.Escape("{type}"), @"(?<type>[0-9]+(?:\.[0-9]+)*)");
            var term = new Regex("^(?:" + token + ")$", RegexOptions.CultureInvariant);
            int balance = 0;

            foreach (char c in expression)
            {
                if (c == '(')
                    balance++;

                if (c == ')')
                    balance--;

                if (balance < 0)
                    throw new InvalidOperationException("Несогласованные скобки в формуле");
            }

            if (balance != 0)
                throw new InvalidOperationException("Несогласованные скобки в формуле");

            string flat = expression.Replace("(", "").Replace(")", "");
            var parts = flat.Split('+');
            var types = new List<string>();

            foreach (string raw in parts)
            {
                var m = term.Match(raw.Trim());

                if (!m.Success)
                    throw new InvalidOperationException("Состав не распознан: поддерживается сумма обычных параметров по шаблону «" + BaseTemplate + "». Формула: " + formula);

                types.Add(m.Groups["type"].Value);
            }

            if (types.Distinct().Count() != types.Count)
                throw new InvalidOperationException("Тип повторяется в формуле; состав неоднозначен: " + formula);

            return types;
        }


        internal string FinishText(string section, List<string> types)
        {
            return (UseSections && SectionInText ? section + " — " : "") + (types.Count == 1 ? "Отделка типа " : "Отделка типов ") + string.Join(", ", types);
        }


        internal void Validate()
        {
            if (new[]
            {
                Source,
                Target,
                Finish,
                Marker,
                BaseTemplate
            }.Any(string.IsNullOrWhiteSpace))
                throw new InvalidOperationException("Заполните имена параметров, маркер и шаблон обычного параметра.");

            if (Source == Target || Source == Finish || Target == Finish)
                throw new InvalidOperationException("Имена исходного параметра и двух параметров записи должны различаться.");

            if (!UseSum && !Fallback)
                throw new InvalidOperationException("Включите поиск SUM или обычного параметра.");

            foreach (string template in UseSum ? new[]
            {
                BaseTemplate,
                SumTemplate
            }

            : new[]
            {
                BaseTemplate
            })
            {
                if (string.IsNullOrWhiteSpace(template) || Regex.Matches(template, Regex.Escape("{type}")).Count != 1)
                    throw new InvalidOperationException("Каждый шаблон должен содержать ровно один {type}.");

                if (Regex.IsMatch(template.Replace("{type}", "").Replace("{section}", ""), "[{}]"))
                    throw new InvalidOperationException("Разрешены только {type} и {section}.");

                if (UseSections && !template.Contains("{section}"))
                    throw new InvalidOperationException("При использовании секций добавьте {section} в шаблоны глобальных параметров.");

                if (!UseSections && template.Contains("{section}"))
                    throw new InvalidOperationException("Уберите {section} из шаблонов или включите секции.");
            }

            if (UseSum && BaseTemplate == SumTemplate)
                throw new InvalidOperationException("Шаблоны обычного и суммарного параметра должны различаться.");

            if (UseSections && string.IsNullOrWhiteSpace(SectionMarker))
                throw new InvalidOperationException("Укажите маркер секции.");
        }


        private static string FilePath(Document d)
        {
            string key = string.IsNullOrEmpty(d.PathName) ? d.ProjectInformation.UniqueId : d.PathName;
            string hash;

            using (var sha = SHA256.Create())
                hash = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(key))).Replace("-", "");

            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ASML", "Revit2020", "Finish-" + hash + ".json");
        }


        internal static FinishRules Load(Document d)
        {
            if (!File.Exists(FilePath(d)))
                return new FinishRules();

            try
            {
                using (var stream = File.OpenRead(FilePath(d)))
                {
                    var s = (FinishRules)new DataContractJsonSerializer(typeof(FinishRules)).ReadObject(stream);
                    s.Validate();
                    return s;
                }
            }
            catch
            {
                F.MessageBox.Show("Правила отделки не удалось прочитать. Загружены значения по умолчанию.");
                return new FinishRules();
            }
        }


        internal void Save(Document d)
        {
            string p = FilePath(d);
            Directory.CreateDirectory(Path.GetDirectoryName(p));

            using (var stream = File.Create(p))
                new DataContractJsonSerializer(typeof(FinishRules)).WriteObject(stream, this);
        }
    }
}
