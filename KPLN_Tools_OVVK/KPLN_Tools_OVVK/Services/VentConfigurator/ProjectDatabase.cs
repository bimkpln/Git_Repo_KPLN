using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Services.VentConfigurator
{
    // Чтение через SQLite из состава Windows 10/11: дополнительные DLL плагину не нужны.
    // https://www.sqlite.org/c3ref/open.html — SQLITE_OPEN_READONLY не создаёт базу.
    internal static class ProjectDatabase
    {
        private const string Query = "SELECT [Name], [Stage], [Code], [MainPath], " +
            "[RevitServerPath], [RevitServerPath2], [RevitServerPath3], [RevitServerPath4] FROM [Projects];";

        internal static IList<ProjectItem> Read(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("База проектов недоступна. Проверьте диск Z: и права на чтение.\n" + path);

            IntPtr database = IntPtr.Zero;
            IntPtr statement = IntPtr.Zero;
            try
            {
                int result = sqlite3_open_v2(Encoding.UTF8.GetBytes(path + "\0"), out database, 1, IntPtr.Zero);
                Check(result, database);
                Check(sqlite3_busy_timeout(database, 1500), database);
                byte[] sql = Encoding.UTF8.GetBytes(Query + "\0");
                IntPtr tail;
                Check(sqlite3_prepare_v2(database, sql, sql.Length, out statement, out tail), database);
                var projects = new List<ProjectItem>();
                while ((result = sqlite3_step(statement)) == 100) // SQLITE_ROW
                {
                    var paths = new List<string>();
                    for (int column = 3; column < 8; column++)
                    {
                        string value = ReadText(statement, column);
                        if (value.Length != 0)
                            paths.Add(value);
                    }
                    string name = ReadText(statement, 0);
                    string stage = ReadText(statement, 1);
                    projects.Add(new ProjectItem
                    {
                        Name = name.Length == 0 ? UnknownProjectName : name,
                        Stage = stage.Length == 0 ? UnknownProjectName : stage,
                        Code = ReadText(statement, 2),
                        Paths = paths.Distinct(StringComparer.OrdinalIgnoreCase).ToArray()
                    });
                }
                if (result != 101)
                    Check(result, database); // SQLITE_DONE
                return projects.OrderBy(p => p.DisplayName, StringComparer.CurrentCultureIgnoreCase).ToList();
            }
            catch (DllNotFoundException ex)
            {
                throw new InvalidOperationException("Компонент SQLite Windows (winsqlite3.dll) недоступен.", ex);
            }
            finally
            {
                if (statement != IntPtr.Zero)
                    sqlite3_finalize(statement);
                if (database != IntPtr.Zero)
                    sqlite3_close(database);
            }
        }

        private static string ReadText(IntPtr statement, int column)
        {
            IntPtr pointer = sqlite3_column_text(statement, column);
            int length = sqlite3_column_bytes(statement, column);
            if (pointer == IntPtr.Zero || length == 0)
                return string.Empty;
            var bytes = new byte[length];
            Marshal.Copy(pointer, bytes, 0, length);
            return CleanDatabaseValue(Encoding.UTF8.GetString(bytes));
        }

        private static void Check(int result, IntPtr database)
        {
            if (result == 0)
                return;
            IntPtr pointer = database == IntPtr.Zero ? IntPtr.Zero : sqlite3_errmsg(database);
            var bytes = new List<byte>();
            if (pointer != IntPtr.Zero)
                for (int i = 0; Marshal.ReadByte(pointer, i) != 0; i++)
                    bytes.Add(Marshal.ReadByte(pointer, i));
            throw new InvalidOperationException("Не удалось прочитать таблицу Projects. SQLite " + result + ": " +
                Encoding.UTF8.GetString(bytes.ToArray()));
        }

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        private static extern int sqlite3_open_v2(byte[] filename, out IntPtr database, int flags, IntPtr vfs);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        private static extern int sqlite3_busy_timeout(IntPtr database, int milliseconds);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        private static extern int sqlite3_prepare_v2(IntPtr database, byte[] sql, int length, out IntPtr statement, out IntPtr tail);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        private static extern int sqlite3_step(IntPtr statement);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        private static extern IntPtr sqlite3_column_text(IntPtr statement, int column);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        private static extern int sqlite3_column_bytes(IntPtr statement, int column);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        private static extern IntPtr sqlite3_errmsg(IntPtr database);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        private static extern int sqlite3_finalize(IntPtr statement);

        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
        private static extern int sqlite3_close(IntPtr database);
    }
}
