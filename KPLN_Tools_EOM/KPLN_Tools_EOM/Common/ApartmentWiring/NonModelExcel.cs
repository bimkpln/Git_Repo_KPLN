using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace KPLN_Tools_EOM.Common.ApartmentWiring
{
    internal sealed class NonModelJob
    {
        internal string Cabinet, Type, SourceType, Sheet;
        internal int Row;
        internal double Quantity;
        internal bool Estimate;
    }


    internal static class NonModelExcel
    {
        private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";


        private static XDocument Xml(ZipArchive z, string part)
        {
            var e = z.GetEntry(part);

            if (e == null)
                throw new InvalidDataException("Не найдено: " + part);

            using (var s = e.Open())
                return XDocument.Load(s);
        }


        internal static List<NonModelJob> Read(string path)
        {
            using (var stream = File.OpenRead(path))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                var book = Xml(zip, "xl/workbook.xml");
                var rels = Xml(zip, "xl/_rels/workbook.xml.rels");
                var strings = zip.GetEntry("xl/sharedStrings.xml") == null ? new List<string>() : Xml(zip, "xl/sharedStrings.xml").Descendants(S + "si").Select(e => string.Concat(e.Descendants(S + "t").Select(t => t.Value))).ToList();
                var result = new List<NonModelJob>();
                var errors = new List<string>();

                foreach (string sheetName in new[]
                {
                    "Кабели по шкафам",
                    "Трубы по шкафам"
                })
                {
                    var sheet = book.Descendants(S + "sheet").SingleOrDefault(e => (string)e.Attribute("name") == sheetName);

                    if (sheet == null)
                        throw new InvalidDataException("Нет листа «" + sheetName + "».");

                    string rid = (string)sheet.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
                    string target = (string)rels.Root.Elements().Single(e => (string)e.Attribute("Id") == rid).Attribute("Target");
                    string part = Uri.UnescapeDataString(new Uri(new Uri("http://xlsx/xl/workbook.xml"), target).AbsolutePath.TrimStart('/'));

                    foreach (var row in Xml(zip, part).Descendants(S + "row"))
                    {
                        int rn = (int)row.Attribute("r");

                        if (rn < 2)
                            continue;

                        var cells = new Dictionary<string, string>();

                        foreach (var cell in row.Elements(S + "c"))
                        {
                            string col = new string (((string)cell.Attribute("r")).TakeWhile(char.IsLetter).ToArray());

                            if (!new[]
                            {
                                "A",
                                "B",
                                "C",
                                "D",
                                "E"
                            }.Contains(col))
                                continue;

                            string raw = (string)cell.Element(S + "v") ?? "", type = (string)cell.Attribute("t");

                            if (type == "s")
                                raw = strings[int.Parse(raw, CultureInfo.InvariantCulture)];

                            if (type == "inlineStr")
                                raw = string.Concat(cell.Descendants(S + "t").Select(e => e.Value));

                            if (cell.Element(S + "f") != null && raw.Length == 0)
                                errors.Add(sheetName + ", строка " + rn + ": нет вычисленного значения Excel в " + col);

                            cells[col] = raw.Trim();
                        }

                        string code = Get(cells, "A"), tn = Get(cells, "B"), qty = Get(cells, "C");

                        if (code.Length + tn.Length + qty.Length == 0)
                            continue;

                        double quantity;

                        if (code.Length == 0 || tn.Length == 0 || !double.TryParse(qty.Replace(" ", "").Replace("\u00a0", "").Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out quantity) || double.IsNaN(quantity) || double.IsInfinity(quantity) || quantity < 0)
                        {
                            errors.Add(sheetName + ", строка " + rn + ": неверный шкаф, тип или количество");
                            continue;
                        }

                        string flag = Get(cells, "E").ToLowerInvariant();

                        if (flag.Length > 0 && !new[]
                        {
                            "1",
                            "0",
                            "да",
                            "нет",
                            "true",
                            "false",
                            "yes",
                            "no",
                            "истина",
                            "ложь"
                        }.Contains(flag))
                        {
                            errors.Add(sheetName + ", строка " + rn + ": неизвестное значение СМ_Смета");
                            continue;
                        }

                        result.Add(new NonModelJob { Cabinet = code, Type = tn, Quantity = quantity, SourceType = Get(cells, "D"), Sheet = sheetName, Row = rn, Estimate = flag.Length == 0 || new[] { "1", "да", "true", "yes", "истина" }.Contains(flag) });
                    }
                }

                if (errors.Count > 0)
                    throw new InvalidDataException(string.Join(Environment.NewLine, errors));

                if (result.Count == 0)
                    throw new InvalidDataException("В Excel нет строк для расстановки.");

                return result;
            }
        }


        private static string Get(Dictionary<string, string> d, string key)
        {
            return d.ContainsKey(key) ? d[key] : "";
        }
    }
}
