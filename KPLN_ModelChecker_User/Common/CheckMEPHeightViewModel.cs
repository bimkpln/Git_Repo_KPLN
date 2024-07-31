using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Newtonsoft.Json;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// ViewModel для окна параметров помещений.
    /// </summary>
    public class CheckMEPHeightViewModel : INotifyPropertyChanged, IJsonSerializable
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _vmcurrentRoomName;
        private string _vmcurrentRoomDepartmentName;
        private double _vmcurrentRoomMinDistance;
        private double _vmcurrentRoomMinElemElevationForCheck;
        private bool _vmisCheckRun;

        [JsonConstructor]
        public CheckMEPHeightViewModel()
        {
        }

        /// <summary>
        /// Создание класса с модели
        /// </summary>
        public CheckMEPHeightViewModel(Room room)
        {
            VMIsCheckRun = false;
            VMCurrentRoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
            VMCurrentRoomDepartmentName = room.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT).AsString();

            VMCurrentRoomMinElemElevationForCheck = 1500;
            if (VMCurrentRoomName.ToLower().Contains("лк"))
                VMCurrentRoomMinDistance = 2200;
            else
                VMCurrentRoomMinDistance = 2000;
        }

        /// <summary>
        /// Запускать проверку?
        /// </summary>
        public bool VMIsCheckRun
        {
            get => _vmisCheckRun;
            set
            {
                _vmisCheckRun = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Имя помещения АР
        /// </summary>
        public string VMCurrentRoomName
        {
            get => _vmcurrentRoomName;
            set
            {
                _vmcurrentRoomName = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Назначение помещения АР
        /// </summary>
        public string VMCurrentRoomDepartmentName
        {
            get => _vmcurrentRoomDepartmentName;
            set
            {
                _vmcurrentRoomDepartmentName = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Мин отметка, при которой элемент считается с потенциальным нарушением
        /// </summary>
        public double VMCurrentRoomMinElemElevationForCheck
        {
            get => _vmcurrentRoomMinElemElevationForCheck;
            set
            {
                _vmcurrentRoomMinElemElevationForCheck = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Минимальная допустимая высота размещения элементов в данном помещении
        /// </summary>
        public double VMCurrentRoomMinDistance
        {
            get => _vmcurrentRoomMinDistance;
            set
            {
                _vmcurrentRoomMinDistance = value;
                NotifyPropertyChanged();
            }
        }

        public object ToJson()
        {
            return new
            {
                this.VMIsCheckRun,
                this.VMCurrentRoomName,
                this.VMCurrentRoomDepartmentName,
                this.VMCurrentRoomMinElemElevationForCheck,
                this.VMCurrentRoomMinDistance,
        };
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
