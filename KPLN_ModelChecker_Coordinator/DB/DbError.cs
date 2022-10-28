using System;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Coordinator.DB
{
    public class DbError
    {
        public string Name { get; }
        public int Count { get; }
        public DbError(string name, int count)
        {
            Name = name;
            Count = count;
        }
        public override string ToString()
        {
            return string.Join(DbVariables.Separator_SubElement, new string[] { Name, Count.ToString() });
        }
        public static List<DbError> TryParseCollection(string data)
        {
            List<DbError> errors = new List<DbError>();
            foreach (string data_unit in data.Split(new string[] { DbVariables.Separator_Element }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] data_unit_part = data_unit.Split(new string[] { DbVariables.Separator_SubElement }, StringSplitOptions.None);
                if (data_unit_part.Length == 2)
                {
                    try
                    {
                        errors.Add(new DbError(data_unit_part[0], int.Parse(data_unit_part[1], System.Globalization.NumberStyles.Integer)));
                    }
                    catch (Exception) { }
                }
            }
            return errors;
        }
    }
}
