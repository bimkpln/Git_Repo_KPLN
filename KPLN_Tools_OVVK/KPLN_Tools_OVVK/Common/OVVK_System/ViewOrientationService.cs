using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    internal enum SystemViewOrientation
    {
        TopFrontLeft,
        TopFrontRight,
        Front,
    }

    internal sealed class ViewOrientationService
    {
        internal int ChangedCount { get; private set; }
        internal List<string> Warnings { get; } = new List<string>();

        internal static View3D[] GetSelectedViews(Document document, UIDocument uiDocument)
        {
            if (uiDocument == null || !document.Equals(uiDocument.Document))
                return new View3D[0];

            return uiDocument.Selection.GetElementIds().Select(document.GetElement).OfType<View3D>()
                .Where(view => !view.IsTemplate).ToArray();
        }

        internal static XYZ GetForwardDirection(SystemViewOrientation orientation)
        {
            // Спереди: камера смотрит вдоль +Y; слева — вдоль +X; сверху — вдоль -Z.
            switch (orientation)
            {
                case SystemViewOrientation.TopFrontLeft: return new XYZ(1, 1, -1).Normalize();
                case SystemViewOrientation.TopFrontRight: return new XYZ(-1, 1, -1).Normalize();
                case SystemViewOrientation.Front: return XYZ.BasisY;
                default: throw new ArgumentOutOfRangeException(nameof(orientation));
            }
        }

        internal void Execute(Document document, View3D[] views, SystemViewOrientation orientation, bool lockViews)
        {
            ChangedCount = 0;
            Warnings.Clear();
            XYZ direction = GetForwardDirection(orientation);
            if (views.Length == 0)
                return;

            int changed = 0;
            using (var transaction = new Transaction(document, "KPLN: Ориентация выбранных 3D-видов"))
            {
                if (transaction.Start() != TransactionStatus.Started)
                    throw new InvalidOperationException("Не удалось начать изменение ориентации видов.");

                foreach (View3D view in views)
                {
                    try
                    {
                        using (var change = new SubTransaction(document))
                        {
                            if (change.Start() != TransactionStatus.Started)
                                throw new InvalidOperationException("Не удалось начать изменение вида.");
                            if (!view.CanSaveOrientation())
                                throw new InvalidOperationException("Ориентацию этого вида нельзя сохранить. Для стандартного вида {3D} сначала измени имя.");

                            if (view.IsLocked)
                                view.Unlock();
                            view.OrientTo(direction);
                            if (lockViews)
                                view.SaveOrientationAndLock();
                            else
                                view.SaveOrientation();

                            if (change.Commit() != TransactionStatus.Committed)
                                throw new InvalidOperationException("Revit отменил изменение вида.");
                        }
                        changed++;
                    }
                    catch (Autodesk.Revit.Exceptions.RegenerationFailedException) { throw; }
                    catch (Exception ex)
                    {
                        Warnings.Add($"Вид «{view.Name}» (Id {view.Id}) не изменён: {ex.Message}");
                    }
                }

                if (changed == 0)
                {
                    transaction.RollBack();
                    return;
                }
                if (transaction.Commit() != TransactionStatus.Committed)
                    throw new InvalidOperationException("Revit отменил изменение ориентации видов.");
                ChangedCount = changed;
            }
        }
    }
}
