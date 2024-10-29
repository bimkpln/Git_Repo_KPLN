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
            _builder.Check(_builder.Doc);

            // Заполняю уровни с учетом секций по геометрии
            using (Transaction t = new Transaction(_builder.Doc))
            {
                t.Start($"{_builder.DocMainTitle}: Параметры захваток_Geom");

                string format = "{0} из " + _builder.AllElementsCount.ToString() + " элементов обработано";
                Progress_Single pb = new Progress_Single($"KPLN_{_builder.DocMainTitle}: Обработка парамеров захваток по геометрии", format, false);
                pb.SetProggresValues(_builder.AllElementsCount, 0);
                pb.ShowProgress();

                _builder.ExecuteGripParams_ByGeom(pb);
                
                pb.Dispose();

                t.Commit();
            }

            // Заполняю уровни с учетом секций по основанию
            using (Transaction t = new Transaction(_builder.Doc))
            {
                t.Start($"{_builder.DocMainTitle}: Параметры захваток_Host");

                string format = "{0} из " + _builder.AllElementsCount.ToString() + " элементов обработано";
                Progress_Single pb = new Progress_Single($"KPLN_{_builder.DocMainTitle}: Обработка парамтеров захваток по геометрии", format, true);
                pb.SetProggresValues(_builder.AllElementsCount, _builder.PbCounter);
                pb.ShowProgress();

                _builder.ExecuteGripParams_ByHost(pb);
                pb.SetBtn_Ok_Enabled();

                t.Commit();
            }

            _builder.CheckNotExecutedElems();

            TaskDialog taskDialog = new TaskDialog("KPLN: Параметры захваток")
            {
                MainInstruction = $"Параметры захваток заполнены для {_builder.PbCounter} из {_builder.AllElementsCount}!",
                CommonButtons = TaskDialogCommonButtons.Ok
            };

            taskDialog.Show();
        }
    }
}