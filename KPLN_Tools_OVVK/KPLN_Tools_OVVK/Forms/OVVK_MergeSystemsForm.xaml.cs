using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker;
using KPLN_Tools_OVVK.Common.OVVK_System;
using KPLN_Tools_OVVK.Forms.Entities;
using System.Collections.Generic;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools_OVVK.Forms
{
    public partial class OVVK_MergeSystemsForm : Window
    {

        private readonly OVVK_SystemManagerM _model;
        private readonly string _modelName;

        public OVVK_MergeSystemsForm(OVVK_SystemManagerM model)
        {
            _model = model;
            // Тот же источник пути, что и в KPLN_Looker: для хранилища берём центральную модель.
            _modelName = SystemProjectConfig.ModelName(model.CurrentDoc);

            InitializeComponent();
            SysIControll.ItemsSource = SysDataToMerge;

            LoadConfig();

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public ObservableCollection<OVVK_MergeSystem> SysDataToMerge { get; set; } = new ObservableCollection<OVVK_MergeSystem>();

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // При добавлении систем окно растёт в обе стороны от своего центра.
            if (IsLoaded && e.WidthChanged && e.PreviousSize.Width > 0)
                Left -= (e.NewSize.Width - e.PreviousSize.Width) / 2;
        }

        private void LoadConfig()
        {
            try
            {
                Document doc = _model.CurrentDoc;
                var saved = SystemProjectConfig.ReadMerge(doc);
                var models = saved?.Models?.Where(item => item != null && !string.IsNullOrWhiteSpace(item.ModelName)
                    && item.ParameterName == _model.ParameterName).ToArray();
                if (models == null || models.Length == 0)
                    return;

                var current = models.FirstOrDefault(item => string.Equals(item.ModelName, _modelName, StringComparison.OrdinalIgnoreCase));
                if (current != null)
                {
                    RestoreGroups(current);
                    return;
                }

                var available = models.Where(item => item.Groups?.Count > 0).OrderBy(item => item.ModelName).ToArray();
                if (available.Length == 0)
                    return;

                SavedModelNames.ItemsSource = available;
                ModelConfigPicker.Visibility = System.Windows.Visibility.Visible;
            }
            catch (Exception ex)
            {
                ConfigStatus.Text = "Не удалось загрузить настройки: " + ex.Message;
            }
        }

        private void LoadModelConfig_Click(object sender, RoutedEventArgs e)
        {
            var selected = SavedModelNames.SelectedItem as OVVK_MergeSystemsConfigM.ModelSettings;
            if (selected == null)
                return;

            RestoreGroups(selected);
            ModelConfigPicker.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void RestoreGroups(OVVK_MergeSystemsConfigM.ModelSettings saved)
        {
            SysDataToMerge.Clear();
            if (saved.Groups == null || saved.Groups.Count == 0)
                return;

            string[] systems = _model.SystemSumParameters.ToArray();
            int skippedGroups = 0;
            foreach (var group in saved.Groups ?? new List<OVVK_MergeSystemsConfigM.Group>())
            {
                string[] names = group == null ? new string[0]
                    : group.SystemNames?.ToArray() ?? new[] { group.FirstParamName, group.LastParamName };
                if (!OVVK_MergeSystem.IsAllowedCount(names.Length)
                    || names.Any(name => string.IsNullOrWhiteSpace(name) || !systems.Contains(name))
                    || names.Distinct().Count() != names.Length)
                {
                    skippedGroups++;
                    continue;
                }

                var restored = new OVVK_MergeSystem(systems);
                restored.SetSystemNames(names);
                SysDataToMerge.Add(restored);
            }

            ConfigStatus.Text = skippedGroups > 0
                ? $"Настройки загружены. Пропущено групп: {skippedGroups} — системы отсутствуют, не заданы или группа некорректна."
                : $"Загружены настройки модели «{saved.ModelName}».";
        }

        private void SaveConfig_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var settings = new OVVK_MergeSystemsConfigM.ModelSettings
                {
                    ModelName = _modelName,
                    ParameterName = _model.ParameterName,
                    Groups = SysDataToMerge.Select(group => new OVVK_MergeSystemsConfigM.Group
                    {
                        SystemNames = group.SelectedSystemNames.ToList(),
                        FirstParamName = group.FirstParamName,
                        LastParamName = group.LastParamName,
                    }).ToList(),
                };
                Document doc = _model.CurrentDoc;
                SystemProjectConfig.SaveMerge(doc, settings);
                ModelConfigPicker.Visibility = System.Windows.Visibility.Collapsed;
                ConfigStatus.Text = $"Настройки модели «{_modelName}» сохранены.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Не удалось сохранить настройки: " + ex.Message,
                    "KPLN: Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            SysDataToMerge.Clear();
            ConfigStatus.Text = "Группы очищены. Нажми «Сохранить», чтобы сохранить пустой список.";
        }
        
        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        private void AddNewGroup_Click(object sender, RoutedEventArgs e)
        {
            SysDataToMerge.Add(new OVVK_MergeSystem(_model.SystemSumParameters.ToArray()));
        }

        private void Run_Click(object sender, RoutedEventArgs e)
        {
            if (!SysDataToMerge.Any())
                return;

            if (SysDataToMerge.Any(group => !OVVK_MergeSystem.IsAllowedCount(group.Systems.Count)
                || group.SelectedSystemNames.Any(string.IsNullOrWhiteSpace)
                || group.SelectedSystemNames.Distinct().Count() != group.Systems.Count))
            {
                MessageBox.Show(this, $"В каждой группе выбери от {OVVK_MergeSystem.MinSystemCount} до {OVVK_MergeSystem.MaxSystemCount} разных систем.",
                    "KPLN: Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string[] names = SysDataToMerge.SelectMany(group => group.SelectedSystemNames).ToArray();
            if (names.Distinct().Count() != names.Length)
            {
                MessageBox.Show(this, "Одна система может участвовать только в одной группе объединения за запуск.",
                    "KPLN: Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
        }

        private void AddSystem_Click(object sender, RoutedEventArgs e)
        {
            ((sender as Button)?.DataContext as OVVK_MergeSystem)?.AddSystem();
        }

        private void RemoveSystem_Click(object sender, RoutedEventArgs e)
        {
            var part = (sender as Button)?.DataContext as OVVK_MergeSystemPart;
            SysDataToMerge.FirstOrDefault(group => group.Systems.Contains(part))?.RemoveSystem(part);
        }

        private void Del_Click(object sender, RoutedEventArgs e)
        {
            OVVK_MergeSystem ent = (sender as Button).DataContext as OVVK_MergeSystem;
            SysDataToMerge.Remove(ent);
        }
    }
}
