using Autodesk.Revit.DB;
using System;
using System.Linq;

namespace KPLN_ViewsAndLists_Ribbon.Common.Views
{
    /// <summary>Общий перенос подрезки для команды «Копировать подрезку» и создания планов по этажам.</summary>
    public static class ViewCropCopyService
    {
        // Вызывается внутри транзакции команды. При области видимости переносится только оформление подрезки.
        public static void Copy(View source, View target, bool copyShape = true)
        {
            using (ViewCropRegionShapeManager from = source.GetCropRegionShapeManager())
            using (ViewCropRegionShapeManager to = target.GetCropRegionShapeManager())
            {
                // Исходный алгоритм ExtCmdCutCopy: передаём контур напрямую, без сдвига по высоте этажа.
                target.CropBoxActive = true;
                target.CropBoxVisible = source.CropBoxVisible;
                if (copyShape)
                {
                    CurveLoop loop = from.GetCropShape().FirstOrDefault();
                    if (loop == null || !to.IsCropRegionShapeValid(loop))
                        throw new InvalidOperationException("Не удалось перенести контур подрезки вида «" + source.Name + "».");

                    to.SetCropShape(loop);
                }


                // Переносим как границы аннотаций, так и флажок включения их подрезки.
                Parameter sourceAnnotation = source.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE);
                Parameter targetAnnotation = target.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE);
                if (sourceAnnotation != null)
                {
                    int active = sourceAnnotation.AsInteger();
                    if (targetAnnotation != null && !targetAnnotation.IsReadOnly)
                        targetAnnotation.Set(active);

                    if (targetAnnotation == null || targetAnnotation.AsInteger() != active)
                        throw new InvalidOperationException("Не удалось перенести включение подрезки аннотаций. Проверьте шаблон целевого вида.");
                }

                if (from.CanHaveAnnotationCrop && to.CanHaveAnnotationCrop)
                {
                    to.BottomAnnotationCropOffset = from.BottomAnnotationCropOffset;
                    to.LeftAnnotationCropOffset = from.LeftAnnotationCropOffset;
                    to.RightAnnotationCropOffset = from.RightAnnotationCropOffset;
                    to.TopAnnotationCropOffset = from.TopAnnotationCropOffset;
                }

                target.CropBoxActive = source.CropBoxActive;
            }
        }
    }
}
