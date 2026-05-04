using Autodesk.Revit.DB;
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
            Progress_Single pb = null;

            try
            {
                _builder.Prepare();
                _builder.Check(_builder.Doc);

                string format = "{0} из " + _builder.AllElementsCount + " элементов обработано";

                // 1. Заполняю уровни с учетом секций по геометрии
                using (Transaction t = new Transaction(_builder.Doc))
                {
                    t.Start($"{_builder.DocMainTitle}: Параметры захваток_Geom");

                    pb = new Progress_Single(
                        $"KPLN_{_builder.DocMainTitle}: Обработка пар-в захваток по геометрии",
                        format,
                        false);

                    pb.SetProggresValues(_builder.AllElementsCount, 0);
                    pb.ShowProgress();

                    _builder.ExecuteGripParams_ByGeom(pb);

                    t.Commit();

                    pb.Close();
                    pb.Dispose();
                    pb = null;
                }

                // 2. Заполняю уровни с учетом секций по основанию
                using (Transaction t = new Transaction(_builder.Doc))
                {
                    t.Start($"{_builder.DocMainTitle}: Параметры захваток_Host");

                    format = "{0} из " + _builder.AllElementsCount + " элементов обработано";

                    pb = new Progress_Single(
                        $"KPLN_{_builder.DocMainTitle}: Обработка пар-в захваток по основанию",
                        format,
                        false); // ВАЖНО: OK не показываем во время процесса

                    pb.SetProggresValues(_builder.AllElementsCount, _builder.PbCounter);
                    pb.ShowProgress();

                    _builder.ExecuteGripParams_ByHost(pb);

                    t.Commit();

                    pb.Close();
                    pb.Dispose();
                    pb = null;
                }

                _builder.CheckNotExecutedElems();
            }
            catch
            {
                if (pb != null)
                {
                    if (!pb.IsDisposed)
                    {
                        pb.Close();
                        pb.Dispose();
                    }

                    pb = null;
                }

                throw; // ВАЖНО: не throw ex;
            }
        }
    }
}