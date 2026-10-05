using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class AutostartConfigItem : Window, INotifyPropertyChanged
    {
        private DBUser _selectedNotificationUser;

        public event PropertyChangedEventHandler PropertyChanged;

        public AutostartConfigItem(DBModuleAutostart assignment, string configurationName)
        {
            InitializeComponent();

            // Редактируем отдельные значения: отмена не меняет назначение в диспетчере.
            ConfigurationName = configurationName;
            InitializeAutostartSchedule(assignment);
            InitializeNotificationUsers(assignment);
            DataContext = this;
        }

        public string ConfigurationName { get; }

        public DBUser[] NotificationUsers { get; private set; } = new DBUser[0];

        public string NotificationUserHelp { get; private set; }

        public int? NotificationUserId => SelectedNotificationUser?.Id;

        public string[] AutostartTimes { get; } = new[] { "06:00", "18:00" };

        public DateTime? AutostartStartDate { get; set; }

        public string AutostartStartTime { get; set; }

        public string AutostartIntervalHours { get; set; }

        public bool SkipWeekends { get; set; }

        public string LastExportDisplay { get; private set; }

        public DateTime? StartDateUtc => AutostartStartDate.HasValue && AutostartTimes.Contains(AutostartStartTime)
            ? (DateTime?)DateTime.SpecifyKind(AutostartStartDate.Value.Date.Add(TimeSpan.ParseExact(AutostartStartTime, @"hh\:mm", CultureInfo.InvariantCulture)), DateTimeKind.Local).ToUniversalTime()
            : null;

        public int? IntervalHours => int.TryParse(AutostartIntervalHours, out int hours) && hours >= 12 && hours % 12 == 0 ? (int?)hours : null;

        private void OnIntervalDecrease(object sender, RoutedEventArgs e) => ChangeInterval(-12);

        private void OnIntervalIncrease(object sender, RoutedEventArgs e) => ChangeInterval(12);

        private void ChangeInterval(int step)
        {
            long next = IntervalHours.HasValue ? (long)IntervalHours.Value + step : 12;
            if (next < 12 || next > int.MaxValue)
                return;

            AutostartIntervalHours = next.ToString(CultureInfo.InvariantCulture);
            OnPropertyChanged(nameof(AutostartIntervalHours));
        }

        private void InitializeAutostartSchedule(DBModuleAutostart moduleAutostart)
        {
            DateTime? startLocal = moduleAutostart == null ? DateTime.Today.AddHours(6)
                : moduleAutostart.StartDateUtc.HasValue ? DateTime.SpecifyKind(moduleAutostart.StartDateUtc.Value, DateTimeKind.Utc).ToLocalTime() : (DateTime?)null;
            AutostartStartDate = startLocal?.Date;
            AutostartStartTime = startLocal?.ToString("HH:mm", CultureInfo.InvariantCulture) ?? AutostartTimes[0];
            AutostartIntervalHours = moduleAutostart == null ? "24" : moduleAutostart.IntervalHours?.ToString(CultureInfo.InvariantCulture);
            SkipWeekends = moduleAutostart?.SkipWeekends ?? false;
            LastExportDisplay = moduleAutostart?.LastExportDateUtc != null
                ? DateTime.SpecifyKind(moduleAutostart.LastExportDateUtc.Value, DateTimeKind.Utc).ToLocalTime().ToString("dd.MM.yyyy HH:mm:ss")
                : "Ещё не завершалась";
        }

        public DBUser SelectedNotificationUser
        {
            get => _selectedNotificationUser;
            set
            {
                _selectedNotificationUser = value;
                OnPropertyChanged();
            }
        }

        private void InitializeNotificationUsers(DBModuleAutostart moduleAutostart)
        {
            int? savedUserId = moduleAutostart?.NotificationUserId;
            NotificationUserHelp = "Выбери сотрудника, подтверди изменения и сохрани выбранные конфигурации в диспетчере автозапуска. " +
                "Для нового назначения по умолчанию выбран текущий пользователь.";
            try
            {
                IEnumerable<DBUser> users = SQLiteMainService.SQLiteUserServiceInst.GetDBUsers();
                if (users == null)
                    throw new InvalidOperationException("Сервис пользователей не вернул список сотрудников.");

                NotificationUsers = users
                    .Where(user => !user.IsFired)
                    .OrderBy(user => user.Surname)
                    .ThenBy(user => user.Name)
                    .ToArray();

                int? selectedId = savedUserId ?? SQLiteMainService.CurrentDBUser?.Id;
                SelectedNotificationUser = NotificationUsers.FirstOrDefault(user => user.Id == selectedId);
            }
            catch (Exception ex)
            {
                NotificationUserHelp = "Не удалось загрузить пользователей. Закрой и снова открой конфигурацию. " + ex.Message;
            }
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        internal void ValidateSettings()
        {
            if (!StartDateUtc.HasValue || !IntervalHours.HasValue
                || Validation.GetHasError(AutostartDatePicker) || !NotificationUserId.HasValue)
                throw new InvalidOperationException("Укажи дату старта, время " + string.Join(" или ", AutostartTimes)
                    + ", интервал от 12 часов с шагом 12 и получателя уведомления.");
        }

        private void OnBtnOkClick(object sender, RoutedEventArgs e)
        {
            try
            {
                ValidateSettings();
                DialogResult = true;
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(this, ex.Message, "KPLN: автозапуск", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
