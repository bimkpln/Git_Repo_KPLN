using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.GripParam
{
    internal class GripByHost
    {
        /// <summary>
        /// Общие метод записи значения секции для вложенных общих семейств
        /// </summary>
        /// <param name="elems">Коллекция элементов для обработки</param>
        /// <param name="sectionParam">Имя параметра номера секции у заполняемых элементов</param>
        /// <param name="levelParam">Имя параметра номера уровня у заполняемых элементов</param>
        /// <param name="pb">Прогресс-бар для демонстрации динамики заполнения</param>
        /// <returns></returns>
        public bool ExecuteByHostFamily(List<Element> elems, string sectionParam, string levelParam, Progress_Single pb, int pbCount)
        {
            
            foreach (Element elem in elems)
            {
                FamilyInstance instance = elem as FamilyInstance;
                Element hostElem = instance.SuperComponent;
                
                string hostElemSectParamValue = hostElem.LookupParameter(sectionParam).AsString();
                Parameter elemSectParam = elem.LookupParameter(sectionParam);
                // Вложенные семейства могут быть заблочены через формулу, для передачи из родительского
                if (!elemSectParam.IsReadOnly)
                    elemSectParam.Set(hostElemSectParamValue);

                string hostElemLevParamValue = hostElem.LookupParameter(levelParam).AsString();
                Parameter elemLevParam = elem.LookupParameter(sectionParam);
                // Вложенные семейства могут быть заблочены через формулу, для передачи из родительского
                if (!elemLevParam.IsReadOnly)
                    elemLevParam.Set(hostElemLevParamValue);

                pb.Update(++pbCount, "Анализ вложенных элементов");
            }
            return true;
        }

        /// <summary>
        /// Общие метод записи значения секции для изоляции ИОС
        /// </summary>
        /// <param name="elems">Коллекция элементов для обработки</param>
        /// <param name="sectionParam">Имя параметра номера секции у заполняемых элементов</param>
        /// <param name="levelParam">Имя параметра номера уровня у заполняемых элементов</param>
        /// <param name="pb">Прогресс-бар для демонстрации динамики заполнения</param>
        /// <returns></returns>
        public bool ExecuteByElementInsulation(Document doc, List<Element> elems, string sectionParam, string levelParam, Progress_Single pb, int pbCount)
        {

            foreach (Element elem in elems)
            {
                InsulationLiningBase insulationLiningBase = elem as InsulationLiningBase;
                if (insulationLiningBase != null)
                {
                    Element hostElem = doc.GetElement(insulationLiningBase.HostElementId);

                    string hostElemSectParamValue = hostElem.LookupParameter(sectionParam).AsString();
                    elem.LookupParameter(sectionParam).Set(hostElemSectParamValue);

                    string hostElemLevParamValue = hostElem.LookupParameter(levelParam).AsString();
                    elem.LookupParameter(levelParam).Set(hostElemLevParamValue);

                    pb.Update(++pbCount, "Анализ изоляции по основам");
                }
                else
                    throw new Exception($"У изоляции с id:{elem.Id} нет основы. Нужно удалить");
            }
            return true;
        }

        /// <summary>
        /// Общие метод записи значения секции для вложенных общих семейств и вложенных семейств витражей
        /// </summary>
        /// <param name="elems">Коллекция элементов для обработки</param>
        /// <param name="sectionParam">Имя параметра номера секции у заполняемых элементов</param>
        /// <param name="levelParam">Имя параметра номера уровня у заполняемых элементов</param>
        /// <param name="pb">Прогресс-бар для демонстрации динамики заполнения</param>
        /// <returns></returns>
        public bool ExecuteByHostFamily_AR(List<Element> elems, string sectionParam, string levelParam, Progress_Single pb, int pbCount)
        {
            foreach (Element elem in elems)
            {
                Element hostElem = null;
                Wall hostWall = null;
                Type elemType = elem.GetType();
                switch (elemType.Name)
                {
                    // Проброс панелей витража на стену
                    case nameof(Panel):
                        Panel panel = (Panel)elem;
                        hostWall = (Wall)panel.Host;
                        hostElem = hostWall;
                        break;

                    // Проброс импостов витража на стену
                    case nameof(Mullion):
                        Mullion mullion = (Mullion)elem;
                        hostWall = (Wall)mullion.Host;
                        hostElem = hostWall;
                        break;

                    // Проброс на вложенные общие семейства
                    default:
                        FamilyInstance instance = elem as FamilyInstance;
                        hostElem = instance.SuperComponent;
                        break;
                }

                string hostElemSectParamValue = hostElem.LookupParameter(sectionParam).AsString();
                Parameter elemSectParam = elem.LookupParameter(sectionParam);
                // Вложенные семейства могут быть заблочены через формулу, для передачи из родительского
                if (!elemSectParam.IsReadOnly)
                    elemSectParam.Set(hostElemSectParamValue);

                string hostElemLevParamValue = hostElem.LookupParameter(levelParam).AsString();
                Parameter elemLevParam = elem.LookupParameter(sectionParam);
                // Вложенные семейства могут быть заблочены через формулу, для передачи из родительского
                if (!elemLevParam.IsReadOnly)
                    elemLevParam.Set(hostElemLevParamValue);

                pb.Update(++pbCount, "Анализ вложенных элементов");
            }
            return true;
        }
    }
}
