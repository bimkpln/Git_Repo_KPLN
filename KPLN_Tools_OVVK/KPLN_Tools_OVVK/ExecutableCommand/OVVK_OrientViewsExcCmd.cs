using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Loader.Common;
using KPLN_Tools_OVVK.Common.OVVK_System;
using KPLN_Tools_OVVK.Forms.Entities;
using System;
using System.Windows;

namespace KPLN_Tools_OVVK.ExecutableCommand
{
    internal sealed class OVVK_OrientViewsExcCmd : IExecutableCommand
    {
        private readonly OVVK_SystemManagerM _model;
        private readonly Window _owner;

        internal OVVK_OrientViewsExcCmd(OVVK_SystemManagerM model, Window owner)
        {
            _model = model;
            _owner = owner;
        }

        public Result Execute(UIApplication app)
        {
            View3D[] views = ViewOrientationService.GetSelectedViews(_model.CurrentDoc, app.ActiveUIDocument);
            if (views.Length == 0)
            {
                _model.SetOperationMessage("Выдели 3D-виды в диспетчере проекта Revit.");
                return Result.Cancelled;
            }
            if (_model.CurrentDoc.IsReadOnly)
                throw new InvalidOperationException("Документ доступен только для чтения.");

            var left = new ButtonToRunEntity("Сверху / спереди / слева", "Ориентировать по вершине видового куба.");
            var right = new ButtonToRunEntity("Сверху / спереди / справа", "Ориентировать по вершине видового куба.");
            var front = new ButtonToRunEntity("Спереди", "Ориентировать по передней грани видового куба.");
            var form = new ButtonToRun($"Ориентация выбранных 3D-видов: {views.Length}", new[] { left, right, front })
            {
                Owner = _owner,
                Title = "KPLN: Ориентация 3D-видов",
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
            };
            if (form.ShowDialog() != true || form.SelectedButton == null)
                return Result.Cancelled;

            MessageBoxResult answer = MessageBox.Show(_owner,
                $"Заблокировать выбранные 3D-виды после поворота? Видов: {views.Length}.\n\n"
                + "Да — сохранить ориентацию и заблокировать виды.\nНет — сохранить ориентацию и оставить виды разблокированными.",
                "KPLN: Блокировка 3D-видов", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (answer != MessageBoxResult.Yes && answer != MessageBoxResult.No)
                return Result.Cancelled;

            SystemViewOrientation orientation = ReferenceEquals(form.SelectedButton, left)
                ? SystemViewOrientation.TopFrontLeft
                : ReferenceEquals(form.SelectedButton, right) ? SystemViewOrientation.TopFrontRight : SystemViewOrientation.Front;
            var service = new ViewOrientationService();
            service.Execute(_model.CurrentDoc, views, orientation, answer == MessageBoxResult.Yes);
            foreach (string warning in service.Warnings)
                HtmlOutput.Print(warning, MessageType.Warning);

            string result = $"Ориентировано 3D-видов: {service.ChangedCount}. Пропущено: {service.Warnings.Count}.";
            _model.SetOperationMessage(result);
            MessageBox.Show(_owner, result, "KPLN: Ориентация 3D-видов", MessageBoxButton.OK,
                service.Warnings.Count > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
            return service.ChangedCount > 0 ? Result.Succeeded : Result.Cancelled;
        }
    }
}
