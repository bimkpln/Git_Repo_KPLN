using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.WorksetUtil;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ModelChecker_Lib.Forms.Entities
{
    public sealed class WSEntity : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _isSelected;

        public WSEntity(Document doc, Workset ws, Workset[] allWorksets)
        {
            RevitWS = ws;

            RevitWSElemsCount = Util.CountWSElems(doc, ws);
            ReplacementWorksets = allWorksets.Where(w => w.Id != ws.Id).ToArray();
            SelectedReplacementWS = RevitWSElemsCount > 0 ? ReplacementWorksets.FirstOrDefault() : null;
        }

        /// <summary>
        /// Текщий рабочий набор
        /// </summary>
        public Workset RevitWS { get; }

        /// <summary>
        /// Количество элементов, которые принадлежат РН
        /// </summary>
        public int RevitWSElemsCount { get; }

        /// <summary>
        /// Пометка выбора
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Коллекция РН файла, для замены
        /// </summary>
        public Workset[] ReplacementWorksets { get; }

        /// <summary>
        /// Выбранный пользователем РН для замены
        /// </summary>
        public Workset SelectedReplacementWS { get; set; }

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
