using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker;
using KPLN_Tools_OVVK.Common.OVVK_System;
using KPLN_Tools_OVVK.Forms.Entities;
using System.Collections.Generic;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools_OVVK.Forms
{
    public partial class OVVK_MergeSystemsForm : Window
    {
        private readonly OVVK_SystemManagerM _model;

        public OVVK_MergeSystemsForm(OVVK_SystemManagerM model)
        {
            _model = model;

            InitializeComponent();
            SysIControll.ItemsSource = SysDataToMerge;

            LoadConfig();

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public ObservableCollection<OVVK_MergeSystem> SysDataToMerge { get; set; } = new ObservableCollection<OVVK_MergeSystem>();

        private string GetConfigName()
        {
            // Разные модели и параметры хранят свои наборы групп.
            string key = (_model.DocumentKey ?? string.Empty).ToUpperInvariant() + "\n" + _model.ParameterName;
            using (var hash = SHA256.Create())
                return "OVVK_MergeSystems_" + BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(key))).Replace("-", "");
        }

        private void LoadConfig()
        {
            try
            {
                var saved = ConfigService.ReadConfigFile<OVVK_MergeSystemsConfigM>(ConfigType.Local, GetConfigName()) as OVVK_MergeSystemsConfigM;
                if (saved == null
                    || !string.Equals(saved.DocumentKey, _model.DocumentKey, StringComparison.OrdinalIgnoreCase)
                    || saved.ParameterName != _model.ParameterName)
                    return;

                string[] systems = _model.SystemSumParameters.ToArray();
                int skippedGroups = 0;
                foreach (var group in saved.Groups ?? new List<OVVK_MergeSystemsConfigM.Group>())
                {
                    if (group == null
                        || string.IsNullOrWhiteSpace(group.FirstParamName)
                        || string.IsNullOrWhiteSpace(group.LastParamName)
                        || !systems.Contains(group.FirstParamName)
                        || !systems.Contains(group.LastParamName))
                    {
                        skippedGroups++;
                        continue;
                    }

                    SysDataToMerge.Add(new OVVK_MergeSystem(systems)
                    {
                        FirstParamName = group.FirstParamName,
                        LastParamName = group.LastParamName,
                    });
                }

                ConfigStatus.Text = skippedGroups > 0
                    ? $"Настройки загружены. Пропущено групп: {skippedGroups} — системы отсутствуют в модели или не заданы."
                    : "Настройки объединения загружены.";
            }
            catch (Exception ex)
            {
                ConfigStatus.Text = "Не удалось загрузить настройки: " + ex.Message;
            }
        }

        private void SaveConfig_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var config = new OVVK_MergeSystemsConfigM
                {
                    DocumentKey = _model.DocumentKey,
                    ParameterName = _model.ParameterName,
                    Groups = SysDataToMerge.Select(group => new OVVK_MergeSystemsConfigM.Group
                    {
                        FirstParamName = group.FirstParamName,
                        LastParamName = group.LastParamName,
                    }).ToList(),
                };
                ConfigService.SaveConfig<OVVK_MergeSystemsConfigM>(ConfigType.Local, config, GetConfigName());
                ConfigStatus.Text = "Настройки объединения сохранены.";
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

            if (SysDataToMerge.Any(group => string.IsNullOrWhiteSpace(group.FirstParamName)
                || string.IsNullOrWhiteSpace(group.LastParamName)
                || group.FirstParamName == group.LastParamName))
            {
                MessageBox.Show(this, "В каждой группе выбери две разные системы.",
                    "KPLN: Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string[] names = SysDataToMerge.SelectMany(group => new[] { group.FirstParamName, group.LastParamName }).ToArray();
            if (names.Distinct().Count() != names.Length)
            {
                MessageBox.Show(this, "Одна система может участвовать только в одной группе объединения за запуск.",
                    "KPLN: Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
        }

        private void Del_Click(object sender, RoutedEventArgs e)
        {
            OVVK_MergeSystem ent = (sender as Button).DataContext as OVVK_MergeSystem;
            SysDataToMerge.Remove(ent);
        }
    }
}
