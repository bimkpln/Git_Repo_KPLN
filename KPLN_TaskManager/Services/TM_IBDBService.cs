using Dapper;
using KPLN_TaskManager.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace KPLN_TaskManager.Services
{
    /// <summary>
    /// Класс для работы с БД содержащих изображения (множество БД, которые содержать ОДНУ строку, для быстроты загрузки)
    /// </summary>
    internal sealed class TM_IBDBService : DBService
    {
        private const string _connectionMainString = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\TaskManager_ImageBuffer";
        private const string _dbMainTableName = "TaskImageBufferItems";

        /// <summary>
        /// Сгенерировать новый путь для отчёта
        /// </summary>
        public static FileInfo GenerateNewPath_DBForTMIB(TaskItemEntity taskItem) =>
            new FileInfo(Path.Combine(_connectionMainString, $"TaskManager_TI_{taskItem.Id}.db"));

        #region Create
        /// <summary>
        /// Создание БД для хранения данных по изображениям по TaskItemEntity
        /// </summary>
        public static void CreateDbFile_ByTaskItem(TaskItemEntity taskItem)
        {
            string connectionString = $@"Data Source={taskItem.PathToImageBufferDB};Version=3;";
            ExecuteNonQuery(
                connectionString,
                $"CREATE TABLE {_dbMainTableName} " +
                    $"({nameof(TaskEntity_ImageBuffer.Id)} INTEGER PRIMARY KEY, " +
                    $"{nameof(TaskEntity_ImageBuffer.TaskEntityId)} INTEGER, " +
                    $"{nameof(TaskEntity_ImageBuffer.ImageBuffer)} BLOB)");
        }

        /// <summary>
        /// Создать новую строку
        /// </summary>
        internal static void CreateDBTaskEntity_ImageBufferItem(TaskItemEntity taskItem)
        {
            string connectionString = $@"Data Source={taskItem.PathToImageBufferDB};Version=3;";

            string query = $"INSERT INTO {_dbMainTableName} " +
                    $"({nameof(TaskEntity_ImageBuffer.Id)}, " +
                    $"{nameof(TaskEntity_ImageBuffer.TaskEntityId)}, " +
                    $"{nameof(TaskEntity_ImageBuffer.ImageBuffer)}) " +
                    $"VALUES " +
                    $"(@{nameof(TaskEntity_ImageBuffer.Id)}, " +
                    $"@{nameof(TaskEntity_ImageBuffer.TaskEntityId)}, " +
                    $"@{nameof(TaskEntity_ImageBuffer.ImageBuffer)}); ";


            var parameters = new DynamicParameters();
            foreach (var buffer in taskItem.TE_ImageBufferColl)
            {
                parameters.Add($"@{nameof(TaskEntity_ImageBuffer.ImageBuffer)}", (object)buffer.ImageBuffer ?? DBNull.Value, DbType.Binary);
                parameters.Add($"@{nameof(TaskEntity_ImageBuffer.TaskEntityId)}", taskItem.Id, DbType.Int32);
                parameters.Add($"@{nameof(TaskEntity_ImageBuffer.Id)}", buffer.Id, DbType.Int32);
            }

            ExecuteNonQuery(connectionString, query, parameters);
        }

        #endregion

        #region Read
        /// <summary>
        /// Получить TaskEntity_ImageBuffer из БД (по TaskItemEntity)
        /// </summary>
        internal static List<TaskEntity_ImageBuffer> GetEntity_ByEntityId(TaskItemEntity taskItem)
        {
            string connectionString = $@"Data Source={taskItem.PathToImageBufferDB};Version=3;";
            return ExecuteQuery<TaskEntity_ImageBuffer>(
                    connectionString,
                    $"SELECT * FROM {_dbMainTableName};")
                .ToList();
        }
        #endregion

        #region Update
        /// <summary>
        /// Обновить/создать редактируемые поля
        /// </summary>
        internal static void UpSetTaskItem_ByTaskItemEntity(TaskItemEntity taskItem)
        {
            string connectionString = $@"Data Source={taskItem.PathToImageBufferDB};Version=3;";

            // 1. Удаляю
            string deleteQuery = $@"
                DELETE FROM {_dbMainTableName}
                WHERE {nameof(TaskEntity_ImageBuffer.TaskEntityId)} = @TaskEntityId;";

            ExecuteNonQuery(connectionString, deleteQuery, new { TaskEntityId = taskItem.Id });


            // 2. Добавляю данные по новой
            string insertQuery = $@"
                INSERT INTO {_dbMainTableName} 
                    ({nameof(TaskEntity_ImageBuffer.Id)},
                     {nameof(TaskEntity_ImageBuffer.TaskEntityId)},
                     {nameof(TaskEntity_ImageBuffer.ImageBuffer)})
                VALUES
                    (@Id, @TaskEntityId, @ImageBuffer);";


            var parameters = new DynamicParameters();
            foreach (var buffer in taskItem.TE_ImageBufferColl)
            {
                parameters.Add($"@{nameof(TaskEntity_ImageBuffer.ImageBuffer)}", (object)buffer.ImageBuffer ?? DBNull.Value, DbType.Binary);
                parameters.Add($"@{nameof(TaskEntity_ImageBuffer.TaskEntityId)}", buffer.TaskEntityId, DbType.Int32);
                parameters.Add($"@{nameof(TaskEntity_ImageBuffer.Id)}", buffer.Id, DbType.Int32);
               
                ExecuteNonQuery(connectionString, insertQuery, parameters);
            }

        }
        #endregion

        #region Delete
        #endregion
    }
}
