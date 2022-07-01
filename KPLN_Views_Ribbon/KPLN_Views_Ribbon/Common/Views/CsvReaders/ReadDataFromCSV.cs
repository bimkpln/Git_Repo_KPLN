using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Views_Ribbon.Views.CsvReaders
{
    public static class ReadDataFromCSV
    {
        public static List<string[]> Read(string filePath)
        {
            //int rowsCount = 0;
            int colsCount = 0;

            List<string[]> linesList = new List<string[]>();
            string[] lines = System.IO.File.ReadAllLines(filePath, Encoding.Default);

            colsCount = lines[0].Split(';').Length;
            string[] line = new string[colsCount];

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i][0] == '#')
                    continue;

                if (!string.IsNullOrEmpty(lines[i]))
                {
                    line = lines[i].Split(';');
                    linesList.Add(line);
                }
            }
            return linesList;
        }

    }
}
