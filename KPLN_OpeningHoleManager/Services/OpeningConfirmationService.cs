using Autodesk.Revit.DB;
using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using KPLN_OpeningHoleManager.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    internal static class OpeningConfirmationService
    {
        // Тот же отдел, который используется для BIM-функций в Clashes.
        private const int BimSubDepartmentId = 8;

        private static DBUser GetBimUser()
        {
            try
            {
                // Проверяем текущую Windows-учётную запись по БД, а не редактируемое имя Revit.
                DBUser user = SQLiteMainService.SQLiteUserServiceInst.GetCurrentDBUser();
                return user?.SubDepartmentId == BimSubDepartmentId ? user : null;
            }
            catch
            {
                return null;
            }
        }

        internal static bool IsBimUser() => GetBimUser() != null;

        internal static OpeningConfirmationResult ConfirmAll(Document document)
        {
            // Повторная проверка при выполнении обязательна: скрытая кнопка не заменяет проверку доступа.
            DBUser user = GetBimUser() ?? throw new UnauthorizedAccessException(
                "«Подтвердить все» доступно только сотруднику BIM-отдела. Не удалось подтвердить доступ по БД.");

            if (document.IsReadOnly || document.IsModifiable)
                throw new InvalidOperationException("Модель недоступна для записи подтверждений. Завершите текущую операцию и повторите.");

            string author = string.Join(" ", new[] { user.Surname, user.Name }.Where(value => !string.IsNullOrWhiteSpace(value)));
            if (string.IsNullOrWhiteSpace(author))
                author = user.SystemName;

            FamilyInstance[] openings = new FilteredElementCollector(document).OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType().Cast<FamilyInstance>().Where(OpeningHistoryService.IsOpening).ToArray();
            var result = new OpeningConfirmationResult { Total = openings.Length };
            if (openings.Length == 0)
                return result;

            var unchanged = new HashSet<string>();
            var pending = new List<string>();
            var failures = new OpeningHistoryService.CacheFailures();

            using (Transaction transaction = new Transaction(document, "KPLN: Подтвердить все отверстия"))
            {
                try
                {
                    transaction.Start();
                    transaction.SetFailureHandlingOptions(transaction.GetFailureHandlingOptions()
                        .SetFailuresPreprocessor(failures).SetClearAfterRollback(true));

                    foreach (FamilyInstance opening in openings)
                    {
                        string id = opening.UniqueId;
                        try
                        {
                            if (OpeningHistoryService.IsConfirmed(opening))
                            {
                                unchanged.Add(id);
                                result.AlreadyConfirmed++;
                                continue;
                            }

                            if (document.IsWorkshared && WorksharingUtils.GetCheckoutStatus(document, opening.Id) == CheckoutStatus.OwnedByOtherUser)
                            {
                                result.Errors[id] = "Не подтверждено: элемент занят другим пользователем";
                                continue;
                            }

                            using (SubTransaction sub = new SubTransaction(document))
                            {
                                try
                                {
                                    sub.Start();
                                    OpeningHistoryService.Confirm(opening, author);
                                    if (sub.Commit() == TransactionStatus.Committed)
                                        pending.Add(id);
                                    else
                                        result.Errors[id] = "Не подтверждено: Revit отменил запись элемента";
                                }
                                catch
                                {
                                    if (sub.GetStatus() == TransactionStatus.Started)
                                        sub.RollBack();
                                    throw;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            result.Errors[id] = "Не подтверждено: " + KPLN_OpeningHoleManager.Services.OpeningUserMessage.FromException(ex);
                        }
                    }

                    if (pending.Count == 0)
                    {
                        transaction.RollBack();
                        return result;
                    }

                    // Успех учитываем только после фиксации общей транзакции.
                    if (transaction.Commit() == TransactionStatus.Committed)
                        result.Confirmed = pending.Count;
                    else
                        foreach (string id in pending)
                            result.Errors[id] = "Не подтверждено: " + failures.ErrorMessage;
                }
                catch (Exception ex)
                {
                    if (transaction.GetStatus() == TransactionStatus.Started)
                        transaction.RollBack();

                    foreach (FamilyInstance opening in openings)
                    {
                        if (!unchanged.Contains(opening.UniqueId) && !result.Errors.ContainsKey(opening.UniqueId))
                            result.Errors[opening.UniqueId] = "Не подтверждено: " + KPLN_OpeningHoleManager.Services.OpeningUserMessage.FromException(ex);
                    }
                }
            }

            return result;
        }
    }
}
