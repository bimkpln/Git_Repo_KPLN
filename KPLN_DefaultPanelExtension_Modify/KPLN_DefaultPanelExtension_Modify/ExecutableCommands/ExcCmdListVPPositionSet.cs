using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_DefaultPanelExtension_Modify.Forms.Models;
using KPLN_Library_PluginActivityWorker;
using KPLN_Loader.Common;
using System;
using System.Windows;

namespace KPLN_DefaultPanelExtension_Modify.ExecutableCommands
{
    internal class ExcCmdListVPPositionSet : IExecutableCommand
    {
        private readonly Element _selVElem;
        private readonly ListVPPositionCreateM _createMItem;

        public ExcCmdListVPPositionSet(Element selVElem, ListVPPositionCreateM createMItem)
        {
            _selVElem = selVElem;
            _createMItem = createMItem;
        }

        public Result Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null)
                return Result.Cancelled;

            Document doc = uidoc.Document;

            if (_selVElem == null)
            {
                System.Windows.MessageBox.Show(
                    "В выобрке нет видов, которые могут быть перемещены",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return Result.Cancelled;
            }


            // Устанавливаю значения для каждого видового экрана
            using (Transaction trans = new Transaction(doc, "KPLN: Положение вида"))
            {
                trans.Start();

                if (_selVElem is Viewport selVP)
                    selVP.SetBoxCenter(GetBoxCenter(_selVElem));
                else if (_selVElem is ScheduleSheetInstance ssi)
                    ssi.Point = GetBoxCenter(_selVElem);

                trans.Commit();
            }


            // Счетчик факта запуска
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName("Положение вида", ModuleData.ModuleName).ConfigureAwait(false);

            return Result.Succeeded;
        }

        private XYZ GetBoxCenter(Element ve)
        {
            // Данные из видового окна
            XYZ[] currentVEPoints = ListVPPositionCreateM.InsertPointsFromVP(ve);
            XYZ currentVEInsertPnt;
            if (ve is Viewport)
                currentVEInsertPnt = currentVEPoints[(int)AlignMode.Center];
            else
                // ПО факту точка вставки для спек - лево низ, но по формированию коллекции - нужен лево верх и +2,1 мм на границы вида (статично)
                currentVEInsertPnt = new XYZ(currentVEPoints[(int)AlignMode.LeftTop].X + 0.007, currentVEPoints[(int)AlignMode.LeftTop].Y, currentVEPoints[(int)AlignMode.LeftTop].Z);

            // Получаю данные смещения в зависимости от конфигурации
            XYZ resultDelta = null;
            if (_createMItem.SelectedAlign == AlignMode.OrignToOrigin)
            {
                try
                {
#if Debug2020 || Revit2020
#else
                    if (ve is Viewport vp)
                        resultDelta = vp.GetProjectionToSheetTransform().Origin - _createMItem.VPTransOrigin;
                    else
                        throw new Exception("Отправь разработчику: " +
                            "в структуре кода ошибка - в выравнивание по СВН попали спеки, хотя это должно быть заблокировано на уровне кнопки wpf");
#endif
                }
                // У легенд и спек - нет привязки к координатам модели
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    System.Windows.MessageBox.Show(
                        "Свомещение внутренних начал применимо только для видов, у которых есть координаты (планы, разрезы, фасады, 3д-виды).\n " +
                            "Для текущего вида - смещение отменено",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    return currentVEInsertPnt;
                }
                catch (Exception ex)
                {
                    throw ex;
                }

            }
            else 
                resultDelta = currentVEPoints[(int)_createMItem.SelectedAlign] - _createMItem.VPInsertPoint;


            return currentVEInsertPnt - resultDelta;
        }
    }
}
