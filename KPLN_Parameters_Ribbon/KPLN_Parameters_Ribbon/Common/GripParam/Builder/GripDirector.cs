using Autodesk.Revit.DB;
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
            
            using (Transaction t = new Transaction(_builder.Doc))
            {
                t.Start($"{_builder.DocMainTitle}: Параметры захваток");

                Print("Заполнение параметров уровня ↑", KPLN_Loader.Preferences.MessageType.Header);
                bool writeLevelParams = _builder.ExecuteLevelParams();
                if (writeLevelParams)
                {
                    Print("Параметры уровня подготовлены успешно!", KPLN_Loader.Preferences.MessageType.Regular);
                }
                
                Print("Заполнение параметров секции ↑", KPLN_Loader.Preferences.MessageType.Header);
                bool writeSectionParams = _builder.ExecuteSectionParams();
                if (writeSectionParams)
                {
                    Print("Параметры секции подготовлены успешно!", KPLN_Loader.Preferences.MessageType.Regular);
                }

                Print("Параметры заполнены!", KPLN_Loader.Preferences.MessageType.Success);
                t.Commit();
            }
        }
    }
}
