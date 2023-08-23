using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Parameters_Ribbon.Forms;

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

            _builder.CountElements();

            // Заполняю уровни с учетом секций
            using (Transaction t = new Transaction(_builder.Doc))
            {
                t.Start($"{_builder.DocMainTitle}: Параметры захваток");

                int max = _builder.AllElementsCount;
                string format = "{0} из " + max.ToString() + " элементов обработано";

                using (Progress_Single pb = new Progress_Single("KPLN: Обработка парамтеров захваток", format, max))
                {
                    bool writeLevelParams = _builder.ExecuteGripParams(pb);
                }

                t.Commit();
            }

            TaskDialog taskDialog = new TaskDialog("KPLN: Параметры захваток")
            {
                MainInstruction = "Параметры захваток заполнены!",
                CommonButtons = TaskDialogCommonButtons.Ok
            };

            taskDialog.Show();
        }
    }
}
