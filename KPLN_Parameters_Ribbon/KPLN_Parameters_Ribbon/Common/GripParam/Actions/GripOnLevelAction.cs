using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Actions
{
    internal class GripOnLevelAction : IGripAction
    {
        public string Name
        {
            get
            {
                return "Заполнение параметра номер этажа для элементов НА уровне";
            }
        }

        /// <summary>
        /// Конструктор класса GripOnLevelAction
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="elems">Список элементов</param>
        /// <param name="floorNumberParamName">Имя параметра номера этажа</param>
        /// <param name="floorTextPosition">Позиция номера этажа в имени уровня</param>
        /// <param name="splitChar">Символ-разделитель в имени уровня</param>
        public GripOnLevelAction(Document doc, List<Element> elems, string floorNumberParamName, int floorTextPosition, char splitChar)
        {

        }

        public bool Compilte()
        {
            
            return true;
        }
    }
}
