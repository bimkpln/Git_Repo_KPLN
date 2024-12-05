using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Parameters_Ribbon.Forms;
using System;

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

            string format = "{0} из " + _builder.AllElementsCount.ToString() + " элементов обработано";
            Progress_Single pb = new Progress_Single($"KPLN_{_builder.DocMainTitle}: Обработка пар-в захваток по геометрии", format, false);
            try
            {
                // Заполняю уровни с учетом секций по геометрии
                using (Transaction t = new Transaction(_builder.Doc))
                {
                    t.Start($"{_builder.DocMainTitle}: Параметры захваток_Geom");

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

                    format = "{0} из " + _builder.AllElementsCount.ToString() + " элементов обработано";
                    pb = new Progress_Single($"KPLN_{_builder.DocMainTitle}: Обработка пар-в захваток по основанию", format, true);
                    pb.SetProggresValues(_builder.AllElementsCount, _builder.PbCounter);
                    pb.ShowProgress();

                    _builder.ExecuteGripParams_ByHost(pb);
                    pb.SetBtn_Ok_Enabled();

                    t.Commit();
                }
            }
            catch (Exception ex)
            {
                pb.Dispose();
                throw ex;
            }

            _builder.CheckNotExecutedElems();
        }
    }
}