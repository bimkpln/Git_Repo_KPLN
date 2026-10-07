using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    internal sealed class InputRow
    {
        internal int Row;
        internal string Name, Formula;
        internal double? Value;
    }


    internal static class Xlsx
    {
        private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";


        private static XDocument Xml(ZipArchive zip, string name)
        {
            var e = zip.GetEntry(name);

            if (e == null)
                throw new InvalidDataException("Отсутствует часть XLSX: " + name);

            using (var stream = e.Open())
                return XDocument.Load(stream);
        }


        internal static List<InputRow> Read(string path)
        {
            using (var stream = File.OpenRead(path))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                var book = Xml(zip, "xl/workbook.xml");
                var sheet = book.Descendants(S + "sheet").SingleOrDefault(e => (string)e.Attribute("name") == "Глобальные параметры");

                if (sheet == null)
                    throw new InvalidDataException("Не найден лист «Глобальные параметры».");

                var rid = (string)sheet.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
                var rel = Xml(zip, "xl/_rels/workbook.xml.rels").Root.Elements().Single(e => (string)e.Attribute("Id") == rid);
                var target = (string)rel.Attribute("Target");
                var uri = new Uri(new Uri("http://xlsx/xl/workbook.xml"), target);
                string part = Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
                var strings = new List<string>();

                if (zip.GetEntry("xl/sharedStrings.xml") != null)
                    strings = Xml(zip, "xl/sharedStrings.xml").Descendants(S + "si").Select(e => string.Concat(e.Descendants(S + "t").Select(t => t.Value))).ToList();

                var result = new List<InputRow>();
                var seen = new HashSet<string>(StringComparer.Ordinal);
                var errors = new List<string>();

                foreach (var r in Xml(zip, part).Descendants(S + "row"))
                {
                    int number = (int)r.Attribute("r");

                    if (number < 2)
                        continue;

                    var cells = new Dictionary<string, string>();

                    foreach (var cell in r.Elements(S + "c"))
                    {
                        string address = (string)cell.Attribute("r");
                        string col = new string (address.TakeWhile(char.IsLetter).ToArray());

                        if (col != "A" && col != "B" && col != "C")
                            continue;

                        string cellValue = (string)cell.Element(S + "v") ?? "";
                        string type = (string)cell.Attribute("t");

                        if (type == "s")
                            cellValue = strings[int.Parse(cellValue, CultureInfo.InvariantCulture)];
                        else if (type == "inlineStr")
                            cellValue = string.Concat(cell.Descendants(S + "t").Select(e => e.Value));

                        if (cell.Element(S + "f") != null)
                            errors.Add("Строка " + number + ", " + col + ": формулу Revit вводите текстом; вычисляемые ячейки Excel A–C не поддерживаются.");

                        cells[col] = cellValue.Trim();
                    }

                    string name = Get(cells, "A"), raw = Get(cells, "B"), formula = Get(cells, "C").TrimStart('=').Trim();

                    if (name.Length + raw.Length + formula.Length == 0)
                        continue;

                    if (name.Length == 0 || !seen.Add(name))
                    {
                        errors.Add("Строка " + number + ": отсутствует или повторяется имя.");
                        continue;
                    }

                    if ((raw.Length == 0) == (formula.Length == 0))
                    {
                        errors.Add("Строка " + number + ": укажите только значение или только формулу.");
                        continue;
                    }

                    double? value = null;

                    if (raw.Length > 0)
                    {
                        double n;

                        if (!double.TryParse(raw.Replace("\u00a0", "").Replace(" ", "").Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out n) || double.IsNaN(n) || double.IsInfinity(n))
                        {
                            errors.Add("Строка " + number + ": значение не является конечным числом.");
                            continue;
                        }

                        value = n;
                    }

                    result.Add(new InputRow { Row = number, Name = name, Formula = formula, Value = value });
                }

                if (errors.Count > 0)
                    throw new InvalidDataException(string.Join(Environment.NewLine, errors));

                if (result.Count == 0)
                    throw new InvalidDataException("На листе нет параметров.");

                return result;
            }
        }


        private static string Get(Dictionary<string, string> d, string k)
        {
            return d.ContainsKey(k) ? d[k] : "";
        }
    }
}
