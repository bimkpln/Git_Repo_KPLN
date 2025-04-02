using Dapper;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_TaskManager.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Threading;

namespace KPLN_TaskManager.Services
{
    internal static class TaskManagerDBService
    {
        private const string _connectionString = @"Data Source=Z:\\Отдел BIM\\03_Скрипты\\08_Базы данных\\KPLN_TaskManager.db;Version=3;";
        private const string _dbMainTableName = "TaskItems";
        private const string _dbCommentTableName = "TaskItemComments";

        #region Create
        /// <summary>
        /// Создать новую единицу задания
        /// </summary>
        internal static void CreateDBTaskItem(TaskItemEntity taskItemEntity)
        {
            ExecuteNonQuery(
                $"INSERT INTO {_dbMainTableName} " +
                $"({nameof(TaskItemEntity.ProjectId)}, " +
                $"{nameof(TaskItemEntity.CreatedTaskUserId)}, " +
                $"{nameof(TaskItemEntity.CreatedTaskDepartmentId)}, " +
                $"{nameof(TaskItemEntity.TaskHeader)}, " +
                $"{nameof(TaskItemEntity.TaskBody)}, " +
                $"{nameof(TaskItemEntity.ImageBuffer)}, " +
                $"{nameof(TaskItemEntity.ModelName)}, " +
                $"{nameof(TaskItemEntity.ModelViewId)}, " +
                $"{nameof(TaskItemEntity.ElementIds)}, " +
                $"{nameof(TaskItemEntity.TaskStatus)}, " +
                $"{nameof(TaskItemEntity.DelegatedDepartmentId)}, " +
                $"{nameof(TaskItemEntity.CreatedTaskData)}, " +
                $"{nameof(TaskItemEntity.LastChangeData)}) " +
                $"VALUES " +
                $"(@{nameof(TaskItemEntity.ProjectId)}, " +
                $"@{nameof(TaskItemEntity.CreatedTaskUserId)}, " +
                $"@{nameof(TaskItemEntity.CreatedTaskDepartmentId)}, " +
                $"@{nameof(TaskItemEntity.TaskHeader)}, " +
                $"@{nameof(TaskItemEntity.TaskBody)}, " +
                $"@{nameof(TaskItemEntity.ImageBuffer)}, " +
                $"@{nameof(TaskItemEntity.ModelName)}, " +
                $"@{nameof(TaskItemEntity.ModelViewId)}, " +
                $"@{nameof(TaskItemEntity.ElementIds)}, " +
                $"@{nameof(TaskItemEntity.TaskStatus)}, " +
                $"@{nameof(TaskItemEntity.DelegatedDepartmentId)}, " +
                $"@{nameof(TaskItemEntity.CreatedTaskData)}, " +
                $"@{nameof(TaskItemEntity.LastChangeData)});",
            taskItemEntity);
        }

        /// <summary>
        /// Создать новую единицу комментария к заданию
        /// </summary>
        internal static void CreateDBTaskItemComment(TaskItemEntity_Comment taskComment)
        {
            ExecuteNonQuery(
                $"INSERT INTO {_dbCommentTableName} " +
                $"({nameof(TaskItemEntity_Comment.TaskItemEntityId)}, " +
                $"{nameof(TaskItemEntity_Comment.DBUserId)}, " +
                $"{nameof(TaskItemEntity_Comment.Message)}, " +
                $"{nameof(TaskItemEntity_Comment.CreatedMsgData)}) " +
                $"VALUES " +
                $"(@{nameof(TaskItemEntity_Comment.TaskItemEntityId)}, " +
                $"@{nameof(TaskItemEntity_Comment.DBUserId)}, " +
                $"@{nameof(TaskItemEntity_Comment.Message)}, " +
                $"@{nameof(TaskItemEntity_Comment.CreatedMsgData)});",
            taskComment);
        }
        #endregion

        #region Read
        /// <summary>
        /// Получить TaskItemEntity по ID TaskItemEntity
        /// </summary>
        internal static TaskItemEntity GetEntity_ByEntityId(int taskEntityId) =>
            ExecuteQuery<TaskItemEntity>(
                $"SELECT * FROM {_dbMainTableName} " +
                $"WHERE {nameof(TaskItemEntity.Id)}='{taskEntityId}';")
                .FirstOrDefault();

        /// <summary>
        /// Получить коллекцию TaskItemEntity по текущему проекту
        /// </summary>
        internal static IEnumerable<TaskItemEntity> GetEntities_ByDBProject(DBProject dBProject) =>
            ExecuteQuery<TaskItemEntity>(
                $"SELECT * FROM {_dbMainTableName} " +
                $"WHERE {nameof(TaskItemEntity.ProjectId)}='{dBProject.Id}';");

        /// <summary>
        /// Получить коллекцию TaskItemEntity по текущему проекту для текущего BIM-пользователя
        /// </summary>
        internal static IEnumerable<TaskItemEntity> GetEntities_ByDBProjectAndBIMUserSubDepId(DBProject dBProject, int userSubDepId, int coordSubDep) =>
            ExecuteQuery<TaskItemEntity>(
                $"SELECT * FROM {_dbMainTableName} " +
                $"WHERE {nameof(TaskItemEntity.ProjectId)}='{dBProject.Id}'" +
                $"AND ({nameof(TaskItemEntity.CreatedTaskDepartmentId)}='{userSubDepId}' AND {nameof(TaskItemEntity.DelegatedDepartmentId)}='{coordSubDep}')"+
                $"OR ({nameof(TaskItemEntity.CreatedTaskDepartmentId)}='{coordSubDep}' AND {nameof(TaskItemEntity.DelegatedDepartmentId)}='{userSubDepId}');");

        /// <summary>
        /// Получить коллекцию TaskItemEntity по текущему проекту, И по текущему разделу пользователя (ЗИ ИЛИ ЗВ)
        /// И по текущей модели
        /// </summary>
        internal static IEnumerable<TaskItemEntity> GetEntities_ByDBPrjIdAndSubDepIdAndDBDoc(int dBPrjId, int dbUserId) =>
            ExecuteQuery<TaskItemEntity>(
                $"SELECT * FROM {_dbMainTableName} " +
                $"WHERE {nameof(TaskItemEntity.ProjectId)}='{dBPrjId}'" +
                $"AND ({nameof(TaskItemEntity.CreatedTaskDepartmentId)}='{dbUserId}'" +
                $"OR {nameof(TaskItemEntity.DelegatedDepartmentId)}='{dbUserId}');");
            

        /// <summary>
        /// Получить коллекцию TaskItemEntity_Comment по текущему TaskItemEntity
        /// </summary>
        internal static IEnumerable<TaskItemEntity_Comment> GetComments_ByTaskItem(TaskItemEntity taskItemEntity) =>
            ExecuteQuery<TaskItemEntity_Comment>(
                $"SELECT * FROM {_dbCommentTableName} " +
                $"WHERE {nameof(TaskItemEntity_Comment.TaskItemEntityId)}='{taskItemEntity.Id}';")
            .Reverse();
        #endregion

        #region Update
        /// <summary>
        /// Обновить редактируемые поля TaskItemEntity по измененому экземпляру TaskItemEntity
        /// </summary>
        /// <param name="tiEnt"></param>
        internal static void UpdateTaskItem_ByTaskItemEntity(TaskItemEntity tiEnt)
        {
            string query = $@"
                UPDATE {_dbMainTableName} 
                SET 
                    {nameof(TaskItemEntity.CreatedTaskUserId)} = @CreatedTaskUserId, 
                    {nameof(TaskItemEntity.CreatedTaskDepartmentId)} = @CreatedTaskDepartmentId, 
                    {nameof(TaskItemEntity.TaskHeader)} = @TaskHeader, 
                    {nameof(TaskItemEntity.DelegatedTaskUserId)} = @DelegatedTaskUserId, 
                    {nameof(TaskItemEntity.DelegatedDepartmentId)} = @DelegatedDepartmentId, 
                    {nameof(TaskItemEntity.BitrixParentTaskId)} = @BitrixParentTaskId, 
                    {nameof(TaskItemEntity.TaskBody)} = @TaskBody, 
                    {nameof(TaskItemEntity.ImageBuffer)} = @ImageBuffer, 
                    {nameof(TaskItemEntity.ModelName)} = @ModelName, 
                    {nameof(TaskItemEntity.ModelViewId)} = @ModelViewId, 
                    {nameof(TaskItemEntity.ElementIds)} = @ElementIds, 
                    {nameof(TaskItemEntity.BitrixTaskId)} = @BitrixTaskId,
                    {nameof(TaskItemEntity.TaskStatus)} = @TaskStatus,
                    {nameof(TaskItemEntity.LastChangeData)} = @LastChangeData
                WHERE 
                    {nameof(TaskItemEntity.Id)} = @Id;";

            var parameters = new
            {
                tiEnt.CreatedTaskUserId,
                tiEnt.CreatedTaskDepartmentId,
                tiEnt.TaskHeader,
                tiEnt.DelegatedTaskUserId,
                tiEnt.DelegatedDepartmentId,
                tiEnt.BitrixParentTaskId,
                tiEnt.TaskBody,
                tiEnt.ImageBuffer,
                tiEnt.ModelName,
                tiEnt.ModelViewId,
                tiEnt.ElementIds,
                tiEnt.BitrixTaskId,
                tiEnt.TaskStatus,
                tiEnt.LastChangeData,
                tiEnt.Id
            };

            ExecuteNonQuery(query, parameters);
        }
        #endregion


        private static void ExecuteNonQuery(string query, object parameters = null)
        {
            const int maxRetries = 3;
            int attempt = 0;
            int timeSleep = 100;

            while (attempt < maxRetries)
            {
                try
                {
                    using (IDbConnection connection = new SQLiteConnection(_connectionString))
                    {
                        connection.Open();
                        connection.Execute(query, parameters);
                        return;
                    }
                }
                catch (SQLiteException ex) when (ex.ErrorCode == (int)SQLiteErrorCode.Busy)
                {
                    attempt++;
                    if (attempt >= maxRetries)
                    {
                        ShowDialog("[KPLN]: Ошибка работы с БД", $"База данных занята. Попытки выполнить запрос ({maxRetries} раза по {timeSleep / 1000.0} с) исчерпаны.");
                        return;
                    }

                    Thread.Sleep(timeSleep);
                }
                catch (Exception ex)
                {
                    ShowDialog("[KPLN]: Ошибка работы с БД", ex.Message);
                    return;
                }
            }
        }

        private static IEnumerable<T> ExecuteQuery<T>(string query, object parameters = null)
        {
            const int maxRetries = 3;
            int attempt = 0;
            int timeSleep = 100;

            while (attempt < maxRetries)
            {
                try
                {
                    using (IDbConnection connection = new SQLiteConnection(_connectionString))
                    {
                        connection.Open();
                        return connection.Query<T>(query, parameters);
                    }
                }
                catch (SQLiteException ex) when (ex.ErrorCode == (int)SQLiteErrorCode.Busy)
                {
                    attempt++;
                    if (attempt >= maxRetries)
                    {
                        ShowDialog("[KPLN]: Ошибка работы с БД", $"База данных занята. Попытки выполнить запрос ({maxRetries} раза по {timeSleep / 1000.0} с) исчерпаны.");
                        return null;
                    }

                    Thread.Sleep(timeSleep);
                }
                catch (Exception ex)
                {
                    ShowDialog("[KPLN]: Ошибка работы с БД", ex.Message);
                    return null;
                }
            }

            return null;
        }

        /// <summary>
        /// Кастомное окно - для возможности вывода информации из другого потока (если использовать Revit API - оно просто не выведется)
        /// </summary>
        private static void ShowDialog(string title, string text)
        {
            System.Windows.Forms.Form form = new System.Windows.Forms.Form()
            {
                Text = title,
                TopMost = true,
                ShowIcon = false,
                StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                AutoSize = true,
                AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink,
                MinimumSize = new System.Drawing.Size(350, 150),
                MaximumSize = new System.Drawing.Size(450, 450),
            };

            System.Windows.Forms.Label textLabel = new System.Windows.Forms.Label()
            {
                Text = text,
                Font = new System.Drawing.Font("GOST Common", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204))),
                AutoSize = true,
                Dock = System.Windows.Forms.DockStyle.Fill,
                MaximumSize = new System.Drawing.Size(440, 0),
                Padding = new System.Windows.Forms.Padding(5),
            };
            form.Controls.Add(textLabel);

            System.Windows.Forms.Button confirmation = new System.Windows.Forms.Button()
            {
                Text = "Ok",
                Location = new System.Drawing.Point((form.Width - 75) / 2, 80),
                Size = new System.Drawing.Size(75, 25),
                Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right,
            };
            confirmation.Click += (sender, e) => { form.Close(); };
            form.Controls.Add(confirmation);

            form.ShowDialog();
        }
    }
}
