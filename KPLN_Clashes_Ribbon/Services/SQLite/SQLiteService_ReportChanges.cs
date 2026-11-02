using Dapper;
using System;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Threading.Tasks;

namespace KPLN_Clashes_Ribbon.Services.SQLite
{
    public sealed partial class SQLiteService_ReportItemsDB
    {
        private volatile bool _trackingReady;

        public sealed class ReportChange
        {
            public long Revision { get; set; }
            public int ItemId { get; set; }
            public string Actor { get; set; }
            public string SessionId { get; set; }
            public int StatusChanged { get; set; }
            public int CommentsChanged { get; set; }
            public int? StatusId { get; set; }
            public int? DelegatedDepartmentId { get; set; }
        }

        public sealed class ChangeSnapshot
        {
            public long Revision { get; set; }
            public SavedItem[] Items { get; set; } = new SavedItem[0];
            public ReportChange[] Changes { get; set; } = new ReportChange[0];
        }

        // Схема добавляется и к старым базам, Items и существующие данные не перестраиваются.
        private void EnsureChangeTracking()
        {
            if (_trackingReady)
                return;

            using (var connection = SQLiteRetryPolicy.CreateConnection(_dbConnectionPath))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    connection.Execute(@"
CREATE TABLE IF NOT EXISTS ReportRevision (Id INTEGER PRIMARY KEY CHECK(Id=1), Revision INTEGER NOT NULL);
INSERT OR IGNORE INTO ReportRevision VALUES (1,0);
CREATE TABLE IF NOT EXISTS ReportWriterContext (Id INTEGER PRIMARY KEY CHECK(Id=1), Actor TEXT, SessionId TEXT);
INSERT OR IGNORE INTO ReportWriterContext VALUES (1,'','');
CREATE TABLE IF NOT EXISTS ReportChanges (
    Revision INTEGER PRIMARY KEY, ItemId INTEGER NOT NULL, Actor TEXT, SessionId TEXT,
    StatusChanged INTEGER NOT NULL, CommentsChanged INTEGER NOT NULL,
    StatusId INTEGER, DelegatedDepartmentId INTEGER);
CREATE INDEX IF NOT EXISTS IX_ReportChanges_ItemRevision ON ReportChanges(ItemId, Revision);",
                        transaction: transaction, commandTimeout: 1);

                    // Старый журнал сохраняем; подробности прошлых событий неизвестны и остаются NULL.
                    string[] columns = connection.Query<string>("SELECT name FROM pragma_table_info('ReportChanges')",
                        transaction: transaction, commandTimeout: 1).ToArray();
                    bool upgraded = false;
                    foreach (string column in new[] { "StatusId", "DelegatedDepartmentId" })
                    {
                        if (columns.Contains(column, StringComparer.OrdinalIgnoreCase))
                            continue;

                        connection.Execute("ALTER TABLE ReportChanges ADD COLUMN " + column + " INTEGER",
                            transaction: transaction, commandTimeout: 1);
                        upgraded = true;
                    }
                    if (upgraded)
                        connection.Execute("DROP TRIGGER IF EXISTS Kpln_ReportItemsChanged", transaction: transaction, commandTimeout: 1);

                    connection.Execute(@"
CREATE TRIGGER IF NOT EXISTS Kpln_ReportItemsChanged
AFTER UPDATE OF StatusId, DelegatedDepartmentId, Comments ON Items
WHEN OLD.StatusId IS NOT NEW.StatusId OR OLD.DelegatedDepartmentId IS NOT NEW.DelegatedDepartmentId
    OR OLD.Comments IS NOT NEW.Comments
BEGIN
    UPDATE ReportRevision SET Revision=Revision+1 WHERE Id=1;
    INSERT INTO ReportChanges (Revision, ItemId, Actor, SessionId, StatusChanged, CommentsChanged, StatusId, DelegatedDepartmentId)
    SELECT r.Revision, NEW.Id, w.Actor, w.SessionId,
        CASE WHEN OLD.StatusId IS NOT NEW.StatusId OR OLD.DelegatedDepartmentId IS NOT NEW.DelegatedDepartmentId THEN 1 ELSE 0 END,
        CASE WHEN OLD.Comments IS NOT NEW.Comments THEN 1 ELSE 0 END,
        NEW.StatusId, NEW.DelegatedDepartmentId
    FROM ReportRevision r, ReportWriterContext w WHERE r.Id=1 AND w.Id=1;
END;", transaction: transaction, commandTimeout: 1);
                    transaction.Commit();
                }
            }

            _trackingReady = true;
        }

        public Task<ChangeSnapshot> ReadChangesAsync(long afterRevision, bool force = false) =>
            SQLiteRetryPolicy.RunAsync(() => ReadChanges(afterRevision, force));

        public ChangeSnapshot ReadChanges(long afterRevision, bool force = false)
        {
            EnsureChangeTracking();
            using (var connection = SQLiteRetryPolicy.CreateConnection(_dbConnectionPath))
            {
                connection.Open();
                long revision = connection.QuerySingle<long>("SELECT Revision FROM ReportRevision WHERE Id=1", commandTimeout: 1);
                if (!force && revision == afterRevision)
                    return new ChangeSnapshot { Revision = revision };

                // Отложенная read-транзакция: версия, журнал и строки принадлежат одному снимку.
                using (var transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted))
                {
                    revision = connection.QuerySingle<long>("SELECT Revision FROM ReportRevision WHERE Id=1", transaction: transaction, commandTimeout: 1);
                    var parameters = new { After = afterRevision, Revision = revision };
                    var result = new ChangeSnapshot { Revision = revision };
                    if (afterRevision >= 0)
                        result.Changes = connection.Query<ReportChange>(
                            "SELECT Revision, ItemId, Actor, SessionId, StatusChanged, CommentsChanged, StatusId, DelegatedDepartmentId FROM ReportChanges WHERE Revision > @After AND Revision <= @Revision ORDER BY Revision",
                            parameters, transaction, commandTimeout: 1).ToArray();

                    string query = @"SELECT i.Id, i.StatusId, i.DelegatedDepartmentId, i.Comments,
                        COALESCE((SELECT MAX(c.Revision) FROM ReportChanges c WHERE c.ItemId=i.Id),0) AS Revision
                        FROM Items i";
                    if (afterRevision >= 0 && !force)
                        query += " WHERE i.Id IN (SELECT ItemId FROM ReportChanges WHERE Revision > @After AND Revision <= @Revision)";
                    result.Items = connection.Query<SavedItem>(query, parameters, transaction, commandTimeout: 1).ToArray();
                    transaction.Commit();
                    return result;
                }
            }
        }
    }
}
