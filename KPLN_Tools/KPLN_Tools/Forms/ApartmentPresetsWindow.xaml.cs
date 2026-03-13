using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class ApartmentPresetsWindow : Window
    {
        public ApartmentPresetData ResultPresetData { get; private set; }

        public ApartmentPresetsWindow(ApartmentPresetData currentData)
        {
            InitializeComponent();

            var vm = new ApartmentPresetsVm(currentData);
            vm.RequestClose += Vm_RequestClose;
            vm.RequestSave += Vm_RequestSave;

            DataContext = vm;
        }

        private void Vm_RequestClose()
        {
            DialogResult = false;
            Close();
        }

        private void Vm_RequestSave(ApartmentPresetData data)
        {
            ResultPresetData = data;
            DialogResult = true;
            Close();
        }
    }

    internal class ApartmentPresetsVm : INotifyPropertyChanged
    {
        public event Action RequestClose;
        public event Action<ApartmentPresetData> RequestSave;

        private int _wallHeight;
        public int WallHeight
        {
            get { return _wallHeight; }
            set
            {
                if (_wallHeight != value)
                {
                    _wallHeight = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _selectedWallType;
        public string SelectedWallType
        {
            get { return _selectedWallType; }
            set
            {
                if (_selectedWallType != value)
                {
                    _selectedWallType = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _selectedEntryDoor;
        public string SelectedEntryDoor
        {
            get { return _selectedEntryDoor; }
            set
            {
                if (_selectedEntryDoor != value)
                {
                    _selectedEntryDoor = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _selectedBathroomDoor;
        public string SelectedBathroomDoor
        {
            get { return _selectedBathroomDoor; }
            set
            {
                if (_selectedBathroomDoor != value)
                {
                    _selectedBathroomDoor = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _selectedRoomDoor;
        public string SelectedRoomDoor
        {
            get { return _selectedRoomDoor; }
            set
            {
                if (_selectedRoomDoor != value)
                {
                    _selectedRoomDoor = value;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<string> WallTypes { get; private set; }
        public ObservableCollection<string> EntryDoors { get; private set; }
        public ObservableCollection<string> BathroomDoors { get; private set; }
        public ObservableCollection<string> RoomDoors { get; private set; }

        public ICommand SaveCommand { get; private set; }
        public ICommand CloseCommand { get; private set; }

        public ApartmentPresetsVm(ApartmentPresetData currentData)
        {
            WallTypes = new ObservableCollection<string>
            {
                "Не выбрано"
            };

            EntryDoors = new ObservableCollection<string>
            {
                "Не выбрано"
            };

            BathroomDoors = new ObservableCollection<string>
            {
                "Не выбрано"
            };

            RoomDoors = new ObservableCollection<string>
            {
                "Не выбрано"
            };

            WallHeight = currentData != null ? currentData.WallHeight : 3000;
            SelectedWallType = currentData != null ? currentData.WallType : "Не выбрано";
            SelectedEntryDoor = currentData != null ? currentData.EntryDoor : "Не выбрано";
            SelectedBathroomDoor = currentData != null ? currentData.BathroomDoor : "Не выбрано";
            SelectedRoomDoor = currentData != null ? currentData.RoomDoor : "Не выбрано";

            if (string.IsNullOrWhiteSpace(SelectedWallType)) SelectedWallType = "Не выбрано";
            if (string.IsNullOrWhiteSpace(SelectedEntryDoor)) SelectedEntryDoor = "Не выбрано";
            if (string.IsNullOrWhiteSpace(SelectedBathroomDoor)) SelectedBathroomDoor = "Не выбрано";
            if (string.IsNullOrWhiteSpace(SelectedRoomDoor)) SelectedRoomDoor = "Не выбрано";

            SaveCommand = new RelayCommand(OnSave);
            CloseCommand = new RelayCommand(OnClose);
        }

        private void OnSave()
        {
            var data = new ApartmentPresetData
            {
                WallHeight = WallHeight,
                WallType = SelectedWallType,
                EntryDoor = SelectedEntryDoor,
                BathroomDoor = SelectedBathroomDoor,
                RoomDoor = SelectedRoomDoor
            };

            RequestSave?.Invoke(data);
        }

        private void OnClose()
        {
            RequestClose?.Invoke();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string p = null)
        {
            var h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(p));
        }
    }
}