using System;
using System.Data.SQLite;
using System.IO;
using System.Reflection;
using KPLN_CoordiantorAI.Common;

internal static class DepartmentPromptTests
{
    private static int _checks;

    private static void Equal<T>(T expected, T actual, string name)
    {
        if (!Equals(expected, actual)) throw new Exception(name + ": expected " + expected + ", got " + actual);
        _checks++;
    }

    private static void Sql(string path, string sql)
    {
        using (var connection = new SQLiteConnection(new SQLiteConnectionStringBuilder
            { DataSource = path, Version = 3 }.ConnectionString))
        {
            connection.Open();
            using (var command = new SQLiteCommand(sql, connection)) command.ExecuteNonQuery();
        }
    }

    public static int Main()
    {
        try
        {
            Equal("AR", DepartmentPromptSettings.ResolveProfile(2), "AR");
            Equal("KR", DepartmentPromptSettings.ResolveProfile(3), "KR");
            foreach (int id in new[] { 4, 5, 6, 7, 41, 51, 71, 72 })
                Equal("IOS", DepartmentPromptSettings.ResolveProfile(id), "IOS " + id);
            foreach (int id in new[] { -1, 0, 1, 8, 999 })
                Equal<string>(null, DepartmentPromptSettings.ResolveProfile(id), "No profile " + id);
            Equal("IOS", DepartmentPromptSettings.ResolveProfile(999, id => id == 999 ? 998 : 4), "New IOS child");
            Equal<string>(null, DepartmentPromptSettings.ResolveProfile(999, id => 999), "Parent cycle");
            Equal<string>(null, DepartmentPromptSettings.ResolveProfile(999, id => id + 1), "Parent depth limit");
            Equal<string>(null, DepartmentPromptSettings.ResolveProfile(8, id => 4), "BIM stays excluded");

            var prompts = new DepartmentPromptSettings { Architecture = "AR-only", Structure = "KR-only", Engineering = "IOS-only" };
            foreach (string profile in new[] { "AR", "KR", "IOS" })
            {
                string context = prompts.BuildContext(profile);
                Equal(true, context.Contains(profile + "-only"), "Selected text " + profile);
                foreach (string other in new[] { "AR", "KR", "IOS" })
                    if (other != profile) Equal(false, context.Contains(other + "-only"), "No mixing " + profile + other);
                Equal(context, prompts.BuildContext(profile), "No accumulated context " + profile);
            }
            Equal("", prompts.BuildContext(null), "Unknown department context");
            var copy = prompts.Clone();
            copy.SetText("AR", "  ");
            Equal("", copy.BuildContext("AR"), "Whitespace disables prompt");
            Equal("AR-only", prompts.Architecture, "Independent clone");
            copy.SetText("IOS", null);
            Equal("", copy.BuildContext("IOS"), "Null disables prompt");

            // Change only in-memory configuration in this isolated test process.
            // Never call SetDatabaseFilePath: it would overwrite the user's configuration.
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "departments-" + Guid.NewGuid().ToString("N") + ".db");
            typeof(CoordinatorAiDatabaseConfig).GetField("_databaseFilePath", BindingFlags.Static | BindingFlags.NonPublic)
                .SetValue(null, path);
            var repository = new CoordinatorAiRepository();
            repository.EnsureDatabase();
            Equal("", repository.LoadDepartmentPrompts().Architecture, "Old database without keys");
            Sql(path, "INSERT INTO Settings (Key, Value, UpdatedAt) VALUES ('GigaChat.SystemPrompt', 'untouched', 'test');");
            var settings = new ExternalModelSettings
            {
                SystemPrompt = "COMMON",
                DepartmentPrompts = new DepartmentPromptSettings
                {
                    Architecture = "Архитектура\r\n'Кавычки' \"текст\"",
                    Structure = "Конструкции",
                    Engineering = "Все инженерные подразделения"
                }
            };
            repository.SaveExternalModelSettings(settings);
            var loaded = repository.LoadExternalModelSettings();
            Equal("COMMON", loaded.SystemPrompt, "Common prompt unchanged");
            Equal(settings.DepartmentPrompts.Architecture, loaded.DepartmentPrompts.Architecture, "AR roundtrip");
            Equal(settings.DepartmentPrompts.Structure, loaded.DepartmentPrompts.Structure, "KR roundtrip");
            Equal(settings.DepartmentPrompts.Engineering, loaded.DepartmentPrompts.Engineering, "IOS roundtrip");
            Equal(settings.DepartmentPrompts.Engineering, repository.LoadDepartmentPrompts().Engineering, "Read-only refresh");
            Equal(true, repository.HasExternalModelSettings(), "Existing department settings");
            using (var connection = new SQLiteConnection("Data Source=" + path))
            {
                connection.Open();
                using (var command = new SQLiteCommand("SELECT Value FROM Settings WHERE Key='GigaChat.SystemPrompt'", connection))
                    Equal("untouched", (string)command.ExecuteScalar(), "Unrelated settings unchanged");
            }
            settings.DepartmentPrompts.Engineering = "";
            repository.SaveExternalModelSettings(settings);
            Equal("", repository.LoadDepartmentPrompts().Engineering, "Clearing instruction");
            Sql(path, "DELETE FROM Settings WHERE Key <> 'ExternalModel.DepartmentPrompt.AR';");
            Equal(true, repository.HasExternalModelSettings(), "Only department prompt is a valid setting");
            string missing = path + ".missing";
            bool failed = false;
            try { repository.LoadDepartmentPrompts(missing); }
            catch (FileNotFoundException) { failed = true; }
            Equal(true, failed, "Missing database reported");
            Equal(false, File.Exists(missing), "Read never creates database");
            Console.WriteLine("PASS: " + _checks + " checks. Test database: " + path);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }
}
