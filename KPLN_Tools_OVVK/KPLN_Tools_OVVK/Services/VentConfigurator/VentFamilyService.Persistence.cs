using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Services.VentConfigurator
{
    internal sealed partial class VentFamilyService
    {
        private static Document GetTargetDocument(UIApplication app, OpenDocumentItem project)
        {
            Document target = project == null ? null : project.Document;
            if (target == null || !target.IsValidObject ||
                !app.Application.Documents.Cast<Document>().Any(d => d.Equals(target)))
                throw new InvalidOperationException("Откройте нужный проект в Revit и повторите действие.");
            if (target.IsFamilyDocument || target.IsLinked || target.IsReadOnly || target.IsModifiable)
                throw new InvalidOperationException("Выбранный проект сейчас недоступен для загрузки. Завершите редактирование и повторите действие.");
            return target;
        }

        private static string ValidateOutputPath(UIApplication app, string sourcePath, string outputPath, string configuredPath)
        {
            string path = Path.GetFullPath(outputPath);
            if (!string.Equals(Path.GetExtension(path), ".rfa", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Для копии необходимо выбрать файл с расширением .rfa.");
            if (SamePath(path, sourcePath) || SamePath(path, SourceFamilyPath) || SamePath(path, LiteralSourceFamilyPath))
                throw new InvalidOperationException("Нельзя сохранять копию поверх исходного семейства. Выберите другой файл.");
            // В утверждённой папке проекта по условию используется исходное имя (Об).
            // Для произвольного пути сохраняем прежнюю защиту от выбора оригинала через UNC.
            if (!SamePath(path, configuredPath)
                && string.Equals(Path.GetFileName(path), Path.GetFileName(sourcePath), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Задайте копии другое имя файла, чтобы отличать её от исходного семейства.");
            if (!Directory.Exists(Path.GetDirectoryName(path)))
                throw new DirectoryNotFoundException("Выбранная папка недоступна.");
            if (File.Exists(path) && (File.GetAttributes(path) & FileAttributes.ReadOnly) != 0)
                throw new IOException(SamePath(path, configuredPath)
                    ? "Семейство проекта доступно только для чтения. Проверьте права на запись в файл проекта."
                    : "Файл копии доступен только для чтения. Выберите другое место сохранения.");
            return path;
        }

        private static bool SamePath(string a, string b)
        {
            // Открытые облачные проекты могут иметь PathName вида BIM 360://... .
            // Такие адреса нельзя передавать в Path.GetFullPath как пути Windows.
            if (!Path.IsPathRooted(a) || !Path.IsPathRooted(b))
                return false;
            return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildAndSaveFamily(UIApplication app, string sourcePath, string outputPath, string typeName, string persistedName, bool freshLibraryCopy, InstallationConfiguration configuration, Action<string> setStage, Action<string> setDiagnostics, out FamilyPackage savedFamily, bool deleting = false, PendingProjectImport import = null)
        {
            savedFamily = null;
            // Подготовка на том же томе, что и результат: готовый RFA заменяет старый файл
            // только после успешной транзакции, SaveAs и Close. Исходник лишь читается File.Copy.
            string workDirectory = Path.Combine(Path.GetDirectoryName(outputPath),
                ".KPLN_Ventilation_" + Guid.NewGuid().ToString("N"));
            string inputDirectory = Path.Combine(workDirectory, "input");
            string stagedDirectory = Path.Combine(workDirectory, "result");
            Document familyDoc = null;
            string warning = string.Empty;
            string diagnosticText = string.Empty;
            string outputHash = File.Exists(outputPath) ? FileHash(outputPath) : null;
            import?.CheckDestination(outputPath);
            try
            {
                setStage("Создание временной копии семейства");
                Directory.CreateDirectory(inputDirectory);
                Directory.CreateDirectory(stagedDirectory);
                string inputPath = Path.Combine(inputDirectory, Path.GetFileName(sourcePath));
                string stagedPath = Path.Combine(stagedDirectory, Path.GetFileName(outputPath));
                File.Copy(sourcePath, inputPath, false);
                File.SetAttributes(inputPath, File.GetAttributes(inputPath) & ~FileAttributes.ReadOnly);
                setStage("Открытие временной копии семейства");
                familyDoc = app.Application.OpenDocumentFile(inputPath);
                if (freshLibraryCopy || import?.FromLibrary == true)
                    RemoveSourceExampleTypes(familyDoc, setStage);
                if (!familyDoc.IsFamilyDocument)
                    throw new InvalidOperationException("Исходный файл не является редактируемым семейством Revit.");

                ApplyFamilyMutation(familyDoc, sourcePath, typeName, persistedName, freshLibraryCopy, configuration, deleting, setStage,
                    value => { diagnosticText = value; setDiagnostics(value); });
                setStage("Чтение сохранённых типов");
                var package = ReadFamilyPackage(familyDoc, outputPath, true);

                setStage("Сохранение подготовленного семейства");
                using (var options = new SaveAsOptions())
                {
                    options.OverwriteExistingFile = false;
                    options.MaximumBackups = 1;
                    familyDoc.SaveAs(stagedPath, options);
                }
                setStage("Закрытие подготовленного семейства");
                if (!familyDoc.Close(false))
                    throw new InvalidOperationException("Не удалось закрыть подготовленную копию. Файл результата не перезаписан.");
                familyDoc = null;

                setStage("Запись итогового файла");
                import?.CheckDestination(outputPath);
                if (outputHash != (File.Exists(outputPath) ? FileHash(outputPath) : null))
                    throw new IOException("Файл семейства изменился во время сохранения. Результат не перезаписан; повторите загрузку файла.");
                if (File.Exists(outputPath))
                {
                    // Без небезопасного fallback Delete + Copy: если том не поддерживает Replace,
                    // операция завершится ошибкой, а существующий результат останется на месте.
                    File.Replace(stagedPath, outputPath, Path.Combine(workDirectory, "previous.rfa"));
                }
                else
                    File.Move(stagedPath, outputPath);
                savedFamily = package;
            }
            catch
            {
                // Только после ошибки сохранения и отката, никогда во время ввода в интерфейсе.
                // Сохраняем первоначальную ошибку даже при неполной диагностике вложенных семейств.
                if (!deleting && configuration != null && familyDoc != null && familyDoc.IsValidObject && !familyDoc.IsModifiable)
                {
                    try
                    {
                        setDiagnostics(diagnosticText + "\n\n" + ReadNestedDiagnostics(familyDoc, configuration));
                    }
                    catch (Exception ex) { setDiagnostics(diagnosticText + "\nВложенные зависимости не прочитаны: " + OperationError.Message(ex)); }
                }
                throw;
            }
            finally
            {
                bool canClean = true;
                if (familyDoc != null && familyDoc.IsValidObject)
                {
                    try
                    {
                        canClean = familyDoc.Close(false);
                    }
                    catch { canClean = false; }
                }
                if (canClean && Directory.Exists(workDirectory))
                {
                    try
                    {
                        Directory.Delete(workDirectory, true);
                    }
                    catch (IOException) { warning = "\nНе удалось удалить временную папку: " + workDirectory; }
                    catch (UnauthorizedAccessException) { warning = "\nНе удалось удалить временную папку: " + workDirectory; }
                }
                else if (!canClean)
                    warning = "\nВременный документ не закрыт: " + workDirectory;
            }
            return warning;
        }

        private static void ApplyFamilyMutation(Document familyDoc, string sourcePath, string typeName, string persistedName, bool freshLibraryCopy, InstallationConfiguration configuration, bool deleting, Action<string> setStage, Action<string> setDiagnostics, Action afterCommit = null)
        {
            setStage("Чтение исходного состояния семейства");
            var manager = familyDoc.FamilyManager;
            var plan = deleting ? null : TypeWritePlan.Make(manager.Types.Cast<FamilyType>().Select(t => t.Name), typeName, persistedName, freshLibraryCopy, configuration?.CopySourceTypeName);
            var initialState = FamilyDiagnosticSnapshot.Read(familyDoc);
            string diagnosticText = "Семейство: " + sourcePath + "\n" + (deleting ? "Удаление типа «" + persistedName + "»."
                : "Режим: " + (plan.Create ? "новый тип средствами Revit" : "редактирование существующего типа")
                    + "\nИсходный тип: " + plan.SourceTypeName + "\nИтоговый тип: " + typeName + "\n\n" + RequestedValuesReport(configuration))
                + "\n\n" + initialState.Report("До изменения семейства");
            Action<string> appendDiagnostics = text =>
            {
                diagnosticText += "\n\n" + text;
                setDiagnostics(diagnosticText);
            };
            setDiagnostics(diagnosticText);
            var sourceInputs = deleting ? null : CaptureTypeInputs(manager,
                manager.Types.Cast<FamilyType>().Single(t => t.Name == plan.SourceTypeName));
            var preserved = manager.Types.Cast<FamilyType>()
                .Where(t => deleting ? t.Name != persistedName : plan.Create || t.Name != persistedName)
                .Select(t => new PreservedType { Type = t, Name = t.Name, Inputs = CaptureTypeInputs(manager, t) }).ToList();
            if (!deleting)
                appendDiagnostics(ReadFamilyDependencies(familyDoc, manager.Types.Cast<FamilyType>().Single(t => t.Name == plan.SourceTypeName))
                    .Report(InstallationConfiguration.InterfaceParameterNames()));

            // Исходный тип нужен только для наследования параметров. Не подтверждаем
            // его геометрию отдельной транзакцией: Revit проверяет сразу итоговый тип.
            // Группа позволяет откатить единственную транзакцию, если проверка после Commit не прошла.
            using (var group = new TransactionGroup(familyDoc, deleting ? "Удалить тип вентиляционной установки" : "Сохранить тип вентиляционной установки"))
            {
                if (group.Start() != TransactionStatus.Started)
                    throw new InvalidOperationException("Не удалось начать изменение семейства.");
                try
                {
                    if (deleting)
                        RunFamilyStage(familyDoc, "Удаление типа", () => DeleteFamilyType(manager, persistedName), null, setStage, appendDiagnostics);
                    else
                    {
                        Action verifyOtherValves = null;
                        RunFamilyStage(familyDoc, "Копирование и настройка итогового типа",
                            () =>
                            {
                                setStage("Выбор исходного типа и создание копии");
                                PrepareFamilyType(manager, typeName, persistedName, freshLibraryCopy, configuration.CopySourceTypeName);
                                VerifyTypeInputs(manager.CurrentType, sourceInputs, null, "Наследование исходных значений нарушено");
                                RefreshDimensionModes(manager, configuration);
                                setStage("Применение параметров к итоговому типу");
                                ApplySectionCount(manager, configuration.IntermediateCount);
                                verifyOtherValves = ApplyValveState(manager, configuration.HasValve);
                                foreach (var flag in configuration.BooleanAssignments())
                                    SetIntegerValue(manager, flag.Key, flag.Value);
                                ApplySectionTypes(familyDoc, configuration);
                                ApplyDimensions(manager, configuration);
                                ApplyInfo(manager, configuration.Info);
                            },
                            () =>
                            {
                                VerifyRequestedValues(familyDoc, configuration);
                                verifyOtherValves?.Invoke();
                                // Связи геометрии и соединителей могут пересчитать служебные размеры
                                // даже без формулы у FamilyParameter. Для текущего типа это штатно:
                                // проверяем записанные настройки выше, а не равенство скрытых значений.
                                // Сохранность остальных типов проверяется отдельно ниже.
                            }, setStage, appendDiagnostics);
                        appendDiagnostics("Итоговый тип подтверждён одной транзакцией. Промежуточные типы отдельно не подтверждались.");
                    }
                    appendDiagnostics(FamilyDiagnosticSnapshot.Read(familyDoc).CompareTo(initialState));
                    setStage("Проверка сохранности остальных типов");
                    foreach (var item in preserved)
                    {
                        if (!manager.Types.Cast<FamilyType>().Any(t => t.Name == item.Name))
                            throw new InvalidOperationException("Исчез исходный тип «" + item.Name + "». Изменения отменены.");
                        VerifyTypeInputs(item.Type, item.Inputs, null, "Изменился другой тип «" + item.Name + "»");
                    }
                    if (!deleting)
                        ReadCalculatedDimensions(familyDoc, configuration);
                    setStage("Подтверждение всех изменений семейства");
                    if (group.Assimilate() != TransactionStatus.Committed)
                        throw new InvalidOperationException("Revit не подтвердил изменения семейства.");
                }
                finally
                {
                    if (group.GetStatus() == TransactionStatus.Started)
                        group.RollBack();
                }
            }
            afterCommit?.Invoke();
        }

        private static void RunFamilyStage(Document document, string stage, Action change, Action verify, Action<string> setStage, Action<string> reportFailures = null)
        {
            setStage(stage);
            using (var transaction = new Transaction(document, stage))
            {
                if (transaction.Start() != TransactionStatus.Started)
                    throw new InvalidOperationException("Не удалось начать этап: " + stage + ".");
                var failures = new TransactionFailures(ignoreEmptyBlockDuplicates: true);
                var options = transaction.GetFailureHandlingOptions();
                options.SetFailuresPreprocessor(failures);
                options.SetClearAfterRollback(true);
                options.SetForcedModalHandling(true);
                transaction.SetFailureHandlingOptions(options);
                try
                {
                    change();
                    // Commit пересчитывает сразу итоговую конфигурацию. Отдельных Regenerate для шаблона и пустой копии нет.
                    setStage(stage + " — подтверждение итогового состояния");
                    if (transaction.Commit() != TransactionStatus.Committed)
                        throw new InvalidOperationException(failures.DescribeErrors("Revit отменил этап «" + stage + "». Изменения семейства отменены."));
                }
                finally
                {
                    if (transaction.GetStatus() == TransactionStatus.Started)
                        transaction.RollBack();
                    if (failures.HasMessages)
                        reportFailures?.Invoke(failures.Describe("Сообщения Revit при подтверждении этапа «" + stage + "»:"));
                }
                setStage(stage + " — проверка результата");
                verify?.Invoke();
            }
        }

        private static Document FindOpenFamily(UIApplication app, string path)
        {
            var document = app.Application.Documents.Cast<Document>().FirstOrDefault(d => d.IsValidObject && SamePath(d.PathName, path));
            if (document != null && (!document.IsFamilyDocument || document.IsLinked))
                throw new InvalidOperationException("По этому пути открыт документ, который нельзя редактировать как семейство.");
            return document;
        }

        private FamilyPackage SaveOpenFamily(Document document, string path, string typeName, string persistedName, InstallationConfiguration configuration, bool deleting, Action<string> setStage, Action<string> setDiagnostics, bool applyChanges = true)
        {
            if (document.IsReadOnly || document.IsModifiable)
                throw new InvalidOperationException("Завершите редактирование в Revit: открытое семейство сейчас недоступно для записи.");
            bool applied = false;
            try
            {
                if (applyChanges)
                    ApplyFamilyMutation(document, path, typeName, persistedName, false, configuration, deleting, setStage, setDiagnostics,
                        () => { applied = true; if (!deleting) _owner.RememberOpenTypeIdentity(path, typeName); });
                setStage("Чтение изменённых типов открытого семейства");
                var family = ReadFamilyPackage(document, path, true);
                setStage("Сохранение открытого семейства");
                document.Save(); // Transaction и TransactionGroup к этому моменту закрыты.
                return family;
            }
            catch (Exception ex)
            {
                if (!applied)
                    throw;
                throw new InvalidOperationException("Изменения применены в открытом семействе, но сохранить файл не удалось. "
                    + "Они остаются в документе Revit; повторите сохранение. " + OperationError.Message(ex), ex);
            }
        }

        private static bool DeleteFamilyType(FamilyManager manager, string typeName)
        {
            if (string.IsNullOrWhiteSpace(typeName) || string.Equals(typeName, BaseFamilyTypeName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Шаблонный тип удалять нельзя.");
            var types = manager.Types.Cast<FamilyType>().ToList();
            var target = types.SingleOrDefault(t => string.Equals(t.Name, typeName, StringComparison.Ordinal));
            if (target == null)
                return false; // Повтор после неудачного Save открытого документа: тип уже удалён в памяти.
            if (types.Count <= 1)
                throw new InvalidOperationException("Нельзя удалить последний тип семейства.");
            manager.CurrentType = target;
            manager.DeleteCurrentType();
            return true;
        }

        private static void OpenSavedFamily(UIApplication app, string path, string typeName, Action<string> setDiagnostics)
        {
            // OpenAndActivateDocument активирует и уже открытый документ; повторного Save здесь нет.
            var document = app.OpenAndActivateDocument(path).Document;
            if (!document.IsFamilyDocument)
                throw new InvalidOperationException("Сохранённый файл не является семейством.");
            var manager = document.FamilyManager;
            var initialState = FamilyDiagnosticSnapshot.Read(document);
            string details = "Открытие сохранённого семейства: " + path + "\nЗапрошенный тип: " + typeName
                + "\n\n" + initialState.Report("После открытия, до выбора типа");
            Action<string> appendDiagnostics = text => { details += "\n\n" + text; setDiagnostics(details); };
            setDiagnostics(details);
            var type = manager.Types.Cast<FamilyType>().SingleOrDefault(t => t.Name == typeName);
            if (type == null)
                throw new InvalidOperationException("В открытом семействе нет типа «" + typeName + "».");
            if (manager.CurrentType != null && manager.CurrentType.Name == typeName)
            {
                appendDiagnostics("Запрошенный тип уже текущий. Транзакция выбора типа не выполнялась.");
                return;
            }
            RunFamilyStage(document, "Показать выбранный тип установки", () => manager.CurrentType = type, null, stage => { }, appendDiagnostics);
            appendDiagnostics(FamilyDiagnosticSnapshot.Read(document).CompareTo(initialState));
        }

        private static string LoadSavedType(UIApplication app, Document target, string path, string typeName)
        {
            if (FindOpenFamily(app, path) == null)
                return LoadType(target, path, typeName);
            // Читаем сохранённые байты; несохранённые изменения открытого RFA не переносятся в проект.
            string directory = Path.Combine(Path.GetTempPath(), "KPLN_Ventilation_Load_" + Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(directory);
                string copy = Path.Combine(directory, Path.GetFileName(path));
                File.Copy(path, copy, false);
                return LoadType(target, copy, typeName);
            }
            finally
            {
                try
                {
                    if (Directory.Exists(directory))
                        Directory.Delete(directory, true);
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }

        private static string FileHash(string path)
        {
            using (var stream = File.OpenRead(path))
            using (var hash = System.Security.Cryptography.SHA256.Create())
                return Convert.ToBase64String(hash.ComputeHash(stream));
        }

        private static TypeWritePlan PrepareFamilyType(FamilyManager manager, string typeName, string persistedName, bool freshLibraryCopy, string copySourceTypeName = null)
        {
            var plan = TypeWritePlan.Make(manager.Types.Cast<FamilyType>().Select(t => t.Name), typeName, persistedName, freshLibraryCopy, copySourceTypeName);
            var source = SelectFamilyType(manager, plan.SourceTypeName);
            if (!plan.Create)
            {
                if (source.Name != typeName)
                    manager.RenameCurrentType(typeName);
                return plan;
            }
            var inherited = CaptureTypeInputs(manager, source);
            manager.NewType(typeName);
            // Только проверка штатного наследования. Ни служебные, ни пользовательские
            // значения нового типа вручную из шаблона не пересобираем.
            VerifyTypeInputs(manager.CurrentType, inherited, null, "Новый тип не унаследовал значения исходного типа «" + source.Name + "»");
            return plan;
        }

        private static FamilyType SelectFamilyType(FamilyManager manager, string name)
        {
            var type = manager.Types.Cast<FamilyType>().Single(t => t.Name == name);
            if (manager.CurrentType == null || manager.CurrentType.Name != name)
                manager.CurrentType = type;
            return type;
        }
    }
}
