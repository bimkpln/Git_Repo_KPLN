using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.Common
{
    internal class SharedEntity
    {
        /// <summary>
        /// Id элемента
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// Имя элемента
        /// </summary>
        
        public string Name { get; set; }
        
        /// <summary>
        /// Заголовок элемента
        /// </summary>
        public string Header{ get; set; }
       
        /// <summary>
        /// Описание элемента
        /// </summary>
        public string Description{ get; set; }

        /// <summary>
        /// Статус, в случае генерации для WPFElement
        /// </summary>
        public Status CurrentStatus { get; set; }

        public SharedEntity(int id)
        {
            Id = id;
        }

        public SharedEntity(int id, string name) : this(id)
        {
            Name = name;
        }

        public SharedEntity(int id, string name, string header) : this (id, name)
        {
            Header = header;
        }
      
        public SharedEntity(int id, string name, string header, string description) : this(id, name, header)
        {
            Description = description;
        }

        public SharedEntity(int id, string name, string header, string description, Status status) : this(id, name, header, description)
        {
            CurrentStatus = status;
        }
    }
}
