using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_TaskManager.Common;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_TaskManager.Services
{
    /// <summary>
    /// Класс для работы с основной БД
    /// </summary>
    internal sealed class TMDBService : DBService
    {
        private const string _connectionString = @"Data Source=Z:\\Отдел BIM\\03_Скрипты\\08_Базы данных\\KPLN_TaskManager.db;Version=3;";
        private const string _dbMainTableName = "TaskItems";
        private const string _dbCommentTableName = "TaskItemComments";

        #region Create
        /// <summary>
        /// Создать новую единицу задания
        /// </summary>
        internal static int CreateDBTaskItem(TaskItemEntity taskItemEntity)
        {
            string query = $"INSERT INTO {_dbMainTableName} " +
                    $"({nameof(TaskItemEntity.ProjectId)}, " +
                    $"{nameof(TaskItemEntity.CreatedTaskUserId)}, " +
                    $"{nameof(TaskItemEntity.CreatedTaskDepartmentId)}, " +
                    $"{nameof(TaskItemEntity.TaskHeader)}, " +
                    $"{nameof(TaskItemEntity.TaskBody)}, " +
                    $"{nameof(TaskItemEntity.PathToImageBufferDB)}, " +
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
                    $"@{nameof(TaskItemEntity.PathToImageBufferDB)}, " +
                    $"@{nameof(TaskItemEntity.ModelName)}, " +
                    $"@{nameof(TaskItemEntity.ModelViewId)}, " +
                    $"@{nameof(TaskItemEntity.ElementIds)}, " +
                    $"@{nameof(TaskItemEntity.TaskStatus)}, " +
                    $"@{nameof(TaskItemEntity.DelegatedDepartmentId)}, " +
                    $"@{nameof(TaskItemEntity.CreatedTaskData)}, " +
                    $"@{nameof(TaskItemEntity.LastChangeData)}); " +
                    $"SELECT last_insert_rowid();";

            var parameters = new
            {
                taskItemEntity.ProjectId,
                taskItemEntity.CreatedTaskUserId,
                taskItemEntity.CreatedTaskDepartmentId,
                taskItemEntity.TaskHeader,
                taskItemEntity.TaskBody,
                taskItemEntity.PathToImageBufferDB,
                taskItemEntity.ModelName,
                taskItemEntity.ModelViewId,
                taskItemEntity.ElementIds,
                taskItemEntity.TaskStatus,
                taskItemEntity.DelegatedDepartmentId,
                taskItemEntity.CreatedTaskData,
                taskItemEntity.LastChangeData,
            };

            return ExecuteInsertWithId(
                _connectionString,
                query,
                parameters);
        }

        /// <summary>
        /// Создать новую единицу комментария к заданию
        /// </summary>
        internal static void CreateDBTaskItemComment(TaskItemEntity_Comment taskComment)
        {
            ExecuteNonQuery(
                _connectionString,
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
        /// Получить значение след. ID
        /// </summary>
        internal static int GetNextId() =>
            ExecuteQuery<int>(_connectionString,
                $"SELECT seq + 1 AS next_id FROM SQLITE_SEQUENCE " +
                $"WHERE name='{_dbMainTableName}';")
                .FirstOrDefault();

        /// <summary>
        /// Получить TaskItemEntity по ID TaskItemEntity
        /// </summary>
        internal static TaskItemEntity GetEntity_ByEntityId(int taskEntityId) =>
            ExecuteQuery<TaskItemEntity>(_connectionString,
                $"SELECT * FROM {_dbMainTableName} " +
                $"WHERE {nameof(TaskItemEntity.Id)}='{taskEntityId}';")
                .FirstOrDefault();

        /// <summary>
        /// Получить коллекцию TaskItemEntity по текущему проекту
        /// </summary>
        internal static IEnumerable<TaskItemEntity> GetEntities_ByDBProject(DBProject dBProject) =>
            ExecuteQuery<TaskItemEntity>(_connectionString,
                $"SELECT * FROM {_dbMainTableName} " +
                $"WHERE {nameof(TaskItemEntity.ProjectId)}='{dBProject.Id}';");

        /// <summary>
        /// Получить коллекцию TaskItemEntity по текущему проекту
        /// И по текущей модели
        /// </summary>
        internal static IEnumerable<TaskItemEntity> GetEntities_ByDBPrjId(int dBPrjId) =>
            ExecuteQuery<TaskItemEntity>(_connectionString,
                $"SELECT * FROM {_dbMainTableName} " +
                $"WHERE {nameof(TaskItemEntity.ProjectId)}='{dBPrjId}';");


        /// <summary>
        /// Получить коллекцию TaskItemEntity_Comment по текущему TaskItemEntity
        /// </summary>
        internal static IEnumerable<TaskItemEntity_Comment> GetComments_ByTaskItem(TaskItemEntity taskItemEntity) =>
            ExecuteQuery<TaskItemEntity_Comment>(_connectionString,
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
                    {nameof(TaskItemEntity.PathToImageBufferDB)} = @PathToImageBufferDB, 
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
                tiEnt.PathToImageBufferDB,
                tiEnt.ModelName,
                tiEnt.ModelViewId,
                tiEnt.ElementIds,
                tiEnt.BitrixTaskId,
                tiEnt.TaskStatus,
                tiEnt.LastChangeData,
                tiEnt.Id
            };

            ExecuteNonQuery(_connectionString, query, parameters);
        }
        #endregion
    }
}
