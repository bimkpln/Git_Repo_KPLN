using Autodesk.Revit.DB;
using KPLN_DefaultPanelExtension_Modify.Common;
using KPLN_Library_ConfigWorker.Core;
using Newtonsoft.Json;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_DefaultPanelExtension_Modify.Forms.Models
{
    public enum AlignMode
    {
        LeftTop=0,
        RightTop=1,
        Center=2,
        LeftBottom=3,
        RightBottom=4,
        // Совмещение Orign к Orign (координаты модели)
        OrignToOrigin = 99,
    }

    public sealed class ListVPPositionCreateM : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _configName;
        private AlignMode _selectedAlign;
        private bool _byOriginVisible;

        [JsonConstructor]
        public ListVPPositionCreateM()
        {

        }

        public ListVPPositionCreateM(Element selVElem)
        {
            VPPoints = InsertPointsFromVP(selVElem);
            SelectedAlign = AlignMode.Center;
#if Debug2020 || Revit2020
#else
            if (selVElem is Viewport selVP)
            {
                try
                {
                    VPTransOrigin = selVP.GetProjectionToSheetTransform().Origin;
                    ByOriginVisible = true;
                }
                // Легенды не имеют связи с коорд модели
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    VPTransOrigin = null;
                    ByOriginVisible = false;
                }
            }
            else
            {
                VPTransOrigin = null;
                ByOriginVisible = false;
            }
#endif
        }

        public string ConfigName
        {
            get => _configName;
            set
            {
                _configName = value;
                OnPropertyChanged();
            }
        }

        public AlignMode SelectedAlign
        {
            get => _selectedAlign;
            set
            {
                _selectedAlign = value;
                OnPropertyChanged();

                SetInsertPointByAlign();
            }
        }

        /// <summary>
        /// Коллекция точек видового экрана
        /// </summary>
        public XYZ[] VPPoints { get; set; }

        /// <summary>
        /// Ккординаты отсчёта трансформа координат видового экрана к модели
        /// </summary>
        [JsonConverter(typeof(XyzJsonConverter))]
        public XYZ VPTransOrigin { get; set; }

        /// <summary>
        /// Точка вставки видового экрана
        /// </summary>
        [JsonConverter(typeof(XyzJsonConverter))]
        public XYZ VPInsertPoint { get; set; }

        /// <summary>
        /// Метка видимости опции выравнивания по совмещению внутренних начал
        /// </summary>
        public bool ByOriginVisible
        {
            get => _byOriginVisible;
            set
            {
                _byOriginVisible = value;
                OnPropertyChanged();
            }
        }

        public object ToJson() => new
        {
            this.ConfigName,
            this.SelectedAlign,
            this.VPTransOrigin,
            this.VPInsertPoint,
        };

        /// <summary>
        /// Получить коллекцию точек для видового окна
        /// </summary>
        /// <param name="selVElem"></param>
        /// <returns></returns>
        public static XYZ[] InsertPointsFromVP(Element selVElem)
        {
            XYZ[] result = new XYZ[5];

            if (selVElem is Viewport selVP)
            {
                Outline outline = selVP.GetBoxOutline();
                XYZ center = selVP.GetBoxCenter();

                // Привязка к AlignMode для понимания соответствия
                result[(int)AlignMode.LeftTop] = new XYZ(outline.MinimumPoint.X, outline.MaximumPoint.Y, center.Z);
                result[(int)AlignMode.RightTop] = new XYZ(outline.MaximumPoint.X, outline.MaximumPoint.Y, center.Z);
                result[(int)AlignMode.Center] = new XYZ(center.X, center.Y, center.Z);
                result[(int)AlignMode.LeftBottom] = new XYZ(outline.MinimumPoint.X, outline.MinimumPoint.Y, center.Z);
                result[(int)AlignMode.RightBottom] = new XYZ(outline.MaximumPoint.X, outline.MinimumPoint.Y, center.Z);
            }
            else if (selVElem is ScheduleSheetInstance ssi)
            {
                BoundingBoxXYZ bbox = ssi.get_BoundingBox(selVElem.Document.ActiveView);
                XYZ center = (bbox.Min + bbox.Max)/2;

                // Привязка к AlignMode для понимания соответствия
                result[(int)AlignMode.LeftTop] = new XYZ(bbox.Min.X, bbox.Max.Y, center.Z);
                result[(int)AlignMode.RightTop] = new XYZ(bbox.Max.X, bbox.Max.Y, center.Z);
                result[(int)AlignMode.Center] = new XYZ(center.X, center.Y, center.Z);
                result[(int)AlignMode.LeftBottom] = new XYZ(bbox.Min.X, bbox.Min.Y, center.Z);
                result[(int)AlignMode.RightBottom] = new XYZ(bbox.Max.X, bbox.Min.Y, center.Z);
            }

            return result;
        }

        private void SetInsertPointByAlign()
        {
            // При чтении из JSON точек не будет
            if (VPPoints != null)
            {
                if (SelectedAlign == AlignMode.OrignToOrigin)
                    VPInsertPoint = VPPoints[(int)AlignMode.Center];
                else
                    VPInsertPoint = VPPoints[(int)SelectedAlign];
            }
        }

        private void OnPropertyChanged([CallerMemberName] string prop = "") => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
    }
}
