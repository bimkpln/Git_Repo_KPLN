using System;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Coordinator.DB
{
    public class DbRowData
    {
        public DateTime DateTime { get; }
        public List<DbError> Errors { get; }
        public DbRowData(string datetime, string data)
        {
            DateTime = DateTime.Parse(datetime);
            Errors = DbError.TryParseCollection(data);
        }
        public DbRowData()
        {
            DateTime = DateTime.Now;
            Errors = new List<DbError>();
        }
        public override string ToString()
        {
            List<string> parts = new List<string>();
            foreach (DbError error in Errors)
            {
                parts.Add(error.ToString());
            }
            return string.Join(DbVariables.Separator_Element, parts);
        }
    }
}
