using System;
using System.Linq;
using System.Windows;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;
using KPLN_Tools_OVVK.ExecutableCommand;
using KPLN_Tools_OVVK.Services.VentConfigurator;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Forms.ViewModels
{
    internal sealed partial class OV_VentConfiguratorVM
    {
        private void UpdateFamily()
        {
            if (Model.IsBusy || !CanUpdateProjectFamily())
                return;
            try
            {
                RaiseRequest(new FamilyRequest
                {
                    Kind = RequestKind.UpdateFamily,
                    OutputPath = Model.WorkingFamilyPath,
                    TypeName = Model.CurrentType?.Name,
                    Types = Model.Types.Where(t => !t.IsCreate).Select(t =>
                {
                    t.MarkChanged();
                    return new FamilyTypeItem
                    {
                        Name = t.Name,
                        PersistedName = t.PersistedName,
                        IsDirty = t.IsDirty,
                        Configuration = t.Configuration?.Copy(),
                        LoadError = t.LoadError
                    };
                }).ToList()
                });
            }
            catch (Exception ex)
            {
                SetBusy(false);
                SetOperationError(OperationError.Create("Обновление семейства", ex, null));
            }
        }

        private void DeleteType(FamilyTypeItem item)
        {
            if (Model.IsBusy || item == null || item.IsCreate || !Model.Types.Contains(item))
                return;
            bool persisted = !string.IsNullOrWhiteSpace(item.PersistedName);
            string text = persisted ? "Тип «" + item.PersistedName + "» будет удалён из файла семейства. После удаления файл будет сразу сохранён. Продолжить?" : "Удалить новый тип «" + item.DisplayName + "»? Он ещё не сохранён в семействе.";
            if (MessageBox.Show(_mainWindow, text, "KPLN. Конфигуратор вент. установок. Удаление типа.", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes)
                return;
            if (!persisted)
            {
                RemoveTypeFromList(item);
                return;
            }

            try
            {
                RaiseRequest(new FamilyRequest
                {
                    Kind = RequestKind.DeleteType,
                    TargetType = item,
                    TypeName = item.Name,
                    PersistedName = item.PersistedName,
                    OutputPath = item.SavedPath,
                    FamilyPath = item.SourcePath ?? item.SavedPath
                });
            }
            catch (Exception ex)
            {
                SetBusy(false);
                SetOperationError(OperationError.Create("Удаление типа", ex, null));
            }
        }

        private void QueueRequest(RequestKind kind)
        {
            if (Model.IsBusy || FamilyService.IsClosed)
                return;
            try
            {
                string typeName = null;
                InstallationConfiguration snapshot = null;
                if (kind != RequestKind.LoadSectionCatalog)
                {
                    if (Model.Configuration == null || Model.CurrentType == null)
                        throw new InvalidOperationException("Сначала создайте тип.");
                    var problems = Model.Configuration.Problems();
                    if (problems.Count > 0)
                    {
                        MessageBox.Show(_mainWindow, "Тип «" + Model.CurrentType.Name + "». Исправьте поля:\n\n" + string.Join("\n", problems.Select(p => p.ParameterName + ": " + p.Message)), "KPLN. Параметры установки", MessageBoxButton.OK, MessageBoxImage.Warning);
                        _mainWindow.NavigateToParameter(problems[0].ParameterName);
                        return;
                    }

                    var missing = Model.Configuration.IncompleteFields();
                    if (string.IsNullOrWhiteSpace(_mainWindow.TypeName))
                        missing.Insert(0, "Имя типа");
                    if (missing.Count > 0)
                    {
                        MessageBox.Show(_mainWindow, "Проверьте состав и параметры установки:\n\n" + string.Join("\n", missing.Select(name => "• " + name)) + "\n\nРазмеры должны быть больше нуля; смещения могут быть нулевыми или отрицательными.", "KPLN. Конфигуратор вент. установок. Не заполнены параметры.", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    Model.CurrentType.MarkChanged();
                    if (kind != RequestKind.Save && (Model.CurrentType.IsDirty || string.IsNullOrWhiteSpace(Model.CurrentType.SavedPath) || string.IsNullOrWhiteSpace(Model.CurrentType.PersistedName)))
                    {
                        MessageBox.Show(_mainWindow, "Сначала сохраните тип и изменения кнопкой «Сохранить».", "KPLN. Конфигуратор вент. установок", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    snapshot = Model.Configuration.Copy();
                    snapshot.ValidateForOutput();
                    typeName = VentTypeName.ValidateTypeName(_mainWindow.TypeName);
                    if (Model.Types.Any(t => !t.IsCreate && t != Model.CurrentType && (Model.CurrentType.PersistedName == null || t.PersistedName != null) && string.Equals((t.Name ?? "").Trim(), typeName, StringComparison.OrdinalIgnoreCase)))
                        throw new InvalidOperationException("Тип с таким именем уже есть в списке. Укажите другое имя.");
                }

                RaiseRequest(new FamilyRequest
                {
                    Kind = kind,
                    TypeName = typeName,
                    Configuration = snapshot,
                    FamilyPath = Model.CurrentType?.SourcePath ?? Model.CurrentType?.SavedPath,
                    OutputPath = Model.CurrentType?.SavedPath,
                    PersistedName = Model.CurrentType?.PersistedName,
                    HasAutomaticName = Model.CurrentType != null && Model.CurrentType.HasAutomaticName,
                    IsDirty = Model.CurrentType != null && Model.CurrentType.IsDirty
                });
            }
            catch (Exception ex)
            {
                SetBusy(false);
                SetOperationError(OperationError.Create("Передача запроса Revit", ex, null));
            }
        }

        private void RaiseRequest(FamilyRequest request)
        {
            if (FamilyService.IsClosed)
                throw new InvalidOperationException("Окно конфигуратора уже закрыто.");
            request.ContextKey = Model.FamilyContextKey;
            SetBusy(true);
            SetStatus("Ожидание Revit. Если выполняется другая команда, завершите её.", false);
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ExcCmd_OV_VentConfigurator(FamilyService, request));
        }
    }
}
