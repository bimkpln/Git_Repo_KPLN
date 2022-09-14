using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    /// <summary>
    /// Распорядитель - создаёт объекты, используюя builder'ы
    /// </summary>
    internal class GripDirector
    {
        private readonly AbstrGripBuilder _builder;
        
        public GripDirector(AbstrGripBuilder builder)
        {
            _builder = builder;
        }
        
        /// <summary>
        /// Создать builder и выполнить запись параметров в элементы Ревит
        /// </summary>
        public void BuildWriter()
        {
            _builder.Prepare();

            _builder.Check();

            // Заполняю уровни с учетом секций (пока не реализован поиск уровня с учетом секции - как вариант организовать отдельную сущность, куда будут помещаться элементы с указанием номера секции, потом этажи заполнять с учетом номера секции)
            using (Transaction t = new Transaction(_builder.Doc))
            {
                t.Start($"{_builder.DocMainTitle}: Параметры захваток");
                
                int max = _builder.AllElementsCount;
                string format = "{0} из " + max.ToString() + " элементов обработано";
                
                using (Progress_Single pb = new Progress_Single("KPLN: Обработка парамтеров секции", format, max))
                {
                    //bool writeLevelParams = _builder.ExecuteSectionParams(pb);
                }
                
                using (Progress_Single pb = new Progress_Single("KPLN: Обработка парамтеров уровня", format, max))
                {
                    bool writeLevelParams = _builder.ExecuteLevelParams(pb);
                }
                
                t.Commit();
            }
            
            Print("Параметры захваток заполнены!", KPLN_Loader.Preferences.MessageType.Success);
        }
    }
}
