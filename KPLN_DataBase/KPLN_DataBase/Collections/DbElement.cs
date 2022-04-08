using System;

namespace KPLN_DataBase.Collections
{
    public abstract class DbElement : IDisposable
    {
        public virtual string TableName { get; }
        protected int _id { get; set; }
        public int Id
        {
            get
            {
                return _id;
            }       
        }
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }
}
