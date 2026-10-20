using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_DefaultPanelExtension_Modify.Forms.Models;
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

            if (_selVElem == null || !_selVElem.IsValidObject || !_selVElem.Document.Equals(doc))
            {
                MessageBox.Show(
                    "Исходный вид удалён или активен другой документ. Выбери вид заново.",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return Result.Cancelled;
            }


            // Проверяем совместимость до открытия транзакции.
            if (_createMItem == null)
                return Result.Cancelled;

            XYZ targetPoint;
            if (_createMItem.SelectedAlign == AlignMode.OrignToOrigin)
            {
                if (!(_selVElem is Viewport viewport) || _createMItem.VPTransOrigin == null)
                    return ShowIncompatibleOrigin();

                try
                {
#if Debug2020 || Revit2020
                    return ShowIncompatibleOrigin();
#else
                    XYZ delta = viewport.GetProjectionToSheetTransform().Origin - _createMItem.VPTransOrigin;
                    targetPoint = viewport.GetBoxCenter() - delta;
#endif
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    return ShowIncompatibleOrigin();
                }
            }
            else
            {
                if ((int)_createMItem.SelectedAlign < 0 || (int)_createMItem.SelectedAlign > 4
                    || _createMItem.VPInsertPoint == null)
                {
                    MessageBox.Show("Конфигурация не содержит корректной точки выравнивания. Создай её заново.",
                        "Положение вида", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return Result.Cancelled;
                }

                targetPoint = GetBoxCenter(_selVElem);
            }


            // Устанавливаю значения для каждого видового экрана
            using (Transaction trans = new Transaction(doc, "KPLN: Положение вида"))
            {
                trans.Start();

                if (_selVElem is Viewport selVP)
                    selVP.SetBoxCenter(targetPoint);
                else if (_selVElem is ScheduleSheetInstance ssi)
                    ssi.Point = targetPoint;

                trans.Commit();
            }


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
            XYZ resultDelta = currentVEPoints[(int)_createMItem.SelectedAlign] - _createMItem.VPInsertPoint;


            return currentVEInsertPnt - resultDelta;
        }

        private static Result ShowIncompatibleOrigin()
        {
            MessageBox.Show(
                "Совмещение внутренних начал требует вида с координатами модели и конфигурации с сохранённым началом. " +
                "Для спецификаций, легенд и чертёжных видов выбери выравнивание по центру или углу.",
                "Положение вида", MessageBoxButton.OK, MessageBoxImage.Information);
            return Result.Cancelled;
        }
    }
}
