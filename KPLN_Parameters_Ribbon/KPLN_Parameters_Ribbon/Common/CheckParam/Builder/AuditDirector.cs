using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.CheckParam.Builder
{
    internal class AuditDirector
    {
        /// <summary>
        /// Распорядитель - создаёт объекты, используюя builder'ы
        /// </summary>
        private readonly AbstrAuditBuilder _builder;

        public AuditDirector(AbstrAuditBuilder builder)
        {
            _builder = builder;
        }

        /// <summary>
        /// Создать builder и выполнить проверку записи параметров в элементы Ревит
        /// </summary>
        public void BuildWriter()
        {
            _builder.Prepare();

            _builder.Check();

            // Заполняю уровни с учетом секций
            using (Transaction t = new Transaction(_builder.Doc))
            {
                t.Start($"{_builder.DocMainTitle}: Проверка парамтеров");

                int max = _builder.AllElementsCount;
                string format = "{0} из " + max.ToString() + " элементов проверено";

                using (Progress_Single pb = new Progress_Single("KPLN: Проверка факта заполнения парамтеров", format, false))
                {
                    pb.SetProggresValues(max, 0);
                    pb.ShowProgress();
                    
                    bool writeLevelParams = _builder.ExecuteParamsAudit(pb);
                }

                t.Commit();
            }

            TaskDialog taskDialog = new TaskDialog("KPLN: Проверка параметров");
            taskDialog.MainInstruction = "Факт заполнения параметров данными подтвержден!";
            taskDialog.CommonButtons = TaskDialogCommonButtons.Ok;

            taskDialog.Show();
        }
    }
}
