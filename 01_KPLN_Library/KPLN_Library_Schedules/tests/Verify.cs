using System;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.Serialization;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Xml.Linq;
using KPLN_Library_Schedules.Common.Schedules;
using KPLN_Library_Schedules.Services.Schedules;
using KPLN_Library_Schedules.Forms;
using KPLN_Library_Schedules.Forms.ViewModels;

public static class ScheduleLibraryVerification
{
    static void Check(bool value, string message) { if (!value) throw new Exception(message); }
    static object Internal(string type, string method, params object[] args)
    {
        return typeof(ScheduleManagerVM).Assembly.GetType("KPLN_Library_Schedules.Services.Schedules." + type)
            .GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, args);
    }
    static ScheduleParameterChoice Parameter(string key, string name)
    {
        return new ScheduleParameterChoice { Key = key, Name = name, SuggestedKind = "Текст", ScheduleCount = 1 };
    }
    static LevelScheduleFilter LevelRule()
    {
        var rule = new LevelScheduleFilter { IsExisting = true, Parameter = Parameter("number", "Номер"), Condition = "Начинается с",
            OriginalCondition = "Начинается с", OriginalValue = "2.1", Value = "2.1", Mode = FilterModes.Keep,
            Visibility = "Виден", InitialVisibility = "Виден", InitialParameterKey = "number", InitialKind = "Текст", OriginalIdentity = "source-field" };
        rule.InitialSignature = rule.Signature; rule.InitialFilterSignature = rule.FilterSignature; rule.SourceIndices.Add("source", 1);
        return rule;
    }
    static BatchScheduleSort SortRule(int index)
    {
        var rule = new BatchScheduleSort { IsExisting = true, Parameter = Parameter("purpose", "Назначение"),
            Position = (index+1).ToString(), Direction = "По возрастанию", Visibility = "Скрыт", InitialParameterKey = "purpose" };
        rule.InitialPosition = rule.Position; rule.InitialVisibility = rule.Visibility; rule.InitialSignature = rule.Signature;
        rule.SourceIndices.Add("source", index); return rule;
    }
    public static void Run(string output)
    {
        var rule = LevelRule(); Check(!rule.HasChanges, "Unchanged level rule");
        rule.Visibility = "Скрыт"; Check(rule.HasChanges && !rule.FilterHasChanges, "Visibility-only handling");
        rule.Parameter = Parameter("purpose", "Назначение"); Check(rule.ParameterChanged && rule.FilterHasChanges, "Replacement not detected");
        rule.Mode = FilterModes.Manual; rule.Overrides["L1"] = ""; rule.Overrides["L2"] = "manual";
        var clone = ScheduleLevelFilterService.Clone(rule);
        Check(clone.InitialParameterKey == "number" && clone.Visibility == "Скрыт" && clone.SourceIndices["source"] == 1, "Clone lost metadata");
        Check(clone.Overrides.Count == 2 && clone.Overrides["L1"] == "", "Clone lost manual empty value");
        clone.Parameter = Parameter("purpose", "Назначение"); Check(clone.Overrides.Count == 2, "Rebinding lost manual values");
        clone.Overrides["L2"] = "changed"; Check(rule.Overrides["L2"] == "manual", "Cancel isolation");
        clone.Parameter = Parameter("number", "Номер"); Check(clone.Overrides.Count == 0, "Parameter replacement kept stale values");

        var levelRule = LevelRule();
        levelRule.IsLevelParameter = true;
        levelRule.Mode = FilterModes.Level;
        levelRule.Parameter = Parameter("purpose", "Назначение");
        Check(!levelRule.IsLevelParameter && !levelRule.Modes.Contains(FilterModes.Level)
            && levelRule.Mode == FilterModes.Pattern, "Non-level replacement retained target-level mode");
        var levelClone = ScheduleLevelFilterService.Clone(levelRule);
        levelClone.Parameter = Parameter("number", "Номер");
        Check(levelClone.IsLevelParameter && levelClone.Modes.Contains(FilterModes.Level), "Clone lost original level parameter");
        levelClone.Mode = FilterModes.Level;
        levelClone.Parameter = Parameter("number", "Номер");
        Check(levelClone.Mode == FilterModes.Level, "Same-key rebind lost target-level mode");
        Check(!levelRule.IsLevelParameter, "Clone changed source level metadata");
        levelRule.Parameter = Parameter("number", "Номер");
        Check(levelRule.IsLevelParameter && levelRule.Modes.Contains(FilterModes.Level), "Returning to level parameter lost mode");

        var first = SortRule(0); var second = SortRule(2);
        var changed = (BatchScheduleSort)Internal("ScheduleSortingService", "Clone", first);
        changed.Parameter = Parameter("number", "Номер"); changed.Position = "2"; changed.Visibility = "Виден";
        var added = new BatchScheduleSort { Parameter = Parameter("area", "Площадь"), Position = "3" };
        var xml = (XElement)Internal("ScheduleSortingService", "SaveRules", new[] { changed, second, added }, "source");
        var loaded = (List<BatchScheduleSort>)Internal("ScheduleSortingService", "LoadRules", new List<BatchScheduleSort> { first, second }, "source", XElement.Parse(xml.ToString()));
        Check(loaded.Count == 3 && loaded[0].Parameter.Key == "number" && loaded[0].Position == "2" && loaded[0].Visibility == "Виден", "Sort settings roundtrip");
        Check(loaded[1].InitialParameterKey == "purpose" && loaded[1].SourceIndices["source"] == 2, "Repeated occurrence lost");
        Check(first.Parameter.Key == "purpose", "Load changed source");
        var matches = (List<int[]>)Internal("ScheduleSortOrderPlanner", "MatchCommon", new List<IList<string>> { new[] { "A", "B" }, new[] { "A", "B", "A" } });
        Check(matches.Count == 2 && matches[0][1] == 0 && matches[1][1] == 1, "Repeated common match");
        Render(output + "/constructor.png", false);
        Render(output + "/levels-editor.png", true);
        var report = ScheduleReportForm.Create(new ScheduleExecutionReport());
        report.ShowActivated = false; report.ShowInTaskbar = false; report.Left = -10000;
        try { report.Show(); Check(report.IsLoaded && report.Content != null, "Report failed to load"); } finally { report.Close(); }
        Console.WriteLine("PASS: level replacement/visibility/manual values/cloning; sorting XML roundtrip/duplicates; shared WPF editor in both modes; report window.");
    }
    static ScheduleManagerVM FakeVM(bool levels)
    {
        var vm = (ScheduleManagerVM)FormatterServices.GetUninitializedObject(typeof(ScheduleManagerVM));
        foreach (var field in typeof(ScheduleManagerVM).GetFields(BindingFlags.Instance | BindingFlags.NonPublic))
        {
            if (field.FieldType.IsGenericType && (field.FieldType.GetGenericTypeDefinition() == typeof(System.Collections.ObjectModel.ObservableCollection<>)
                || field.FieldType.GetGenericTypeDefinition() == typeof(List<>) || field.FieldType.GetGenericTypeDefinition() == typeof(Dictionary<,>)))
                field.SetValue(vm, Activator.CreateInstance(field.FieldType));
        }
        typeof(ScheduleManagerVM).GetField("<IsByLevels>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(vm, levels);
        BatchScheduleFilter row = levels ? (BatchScheduleFilter)LevelRule() : new BatchScheduleFilter { IsExisting = true, Parameter = Parameter("number", "Номер"), Condition = "Начинается с", Value = "{Исходное}", Visibility = "Виден" };
        vm.Parameters.Add(row.Parameter); vm.BatchRules.Add(row);
        var sort = SortRule(0); vm.SortRules.Add(sort); vm.SortParameters.Add(sort.Parameter);
        vm.SortPositions.Add("Исходное"); vm.SortPositions.Add("1");
        return vm;
    }
    static void Render(string file, bool levels)
    {
        var vm = FakeVM(levels);
        Window window;
        if (levels)
        {
            // Проверяем раскладку окна без доступа к документу Revit; обработчики проверяются компиляцией.
            var source = System.IO.Path.GetFullPath(System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(typeof(ScheduleManagerVM).Assembly.Location), "../../../Forms/ScheduleFiltersForm.xaml"));
            var xaml = XElement.Parse(System.IO.File.ReadAllText(source));
            foreach (var attribute in xaml.DescendantsAndSelf().Attributes().ToList())
            {
                if (attribute.Name.LocalName == "Class" || attribute.Name.LocalName == "Click" || attribute.Name.LocalName == "BeginningEdit")
                    attribute.Remove();
            }
            string markup = xaml.ToString().Replace("clr-namespace:KPLN_Library_Schedules.Forms\"", "clr-namespace:KPLN_Library_Schedules.Forms;assembly=KPLN_Library_Schedules\"")
                .Replace("clr-namespace:KPLN_Library_Schedules.Common.Schedules\"", "clr-namespace:KPLN_Library_Schedules.Common.Schedules;assembly=KPLN_Library_Schedules\"")
                .Replace("Source=\"ScheduleManagerStyles.xaml\"", "Source=\"/KPLN_Library_Schedules;component/Forms/ScheduleManagerStyles.xaml\"");
            window = (Window)System.Windows.Markup.XamlReader.Parse(markup);
            window.DataContext = new { ScheduleName = "Экспликация помещений — проверка макета", Editor = vm,
                Values = new[] { new LevelFilterValueRow { Level = new ScheduleLevelRow(), Field = "Номер / Начинается с", Mode = FilterModes.Keep, Value = "2.1" } } };
        }
        else window = new ScheduleManagerForm(vm);
        window.ShowActivated = false; window.ShowInTaskbar = false; window.Left = -10000;
        try
        {
            window.Show();
            if (!levels) ((System.Windows.Controls.TabControl)window.FindName("Tabs")).SelectedIndex = 1;
            window.Dispatcher.Invoke(new Action(() => {}), DispatcherPriority.ContextIdle); window.UpdateLayout();
            var visual = (FrameworkElement)window.Content;
            var bitmap = new RenderTargetBitmap((int)visual.ActualWidth, (int)visual.ActualHeight, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(visual); var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap));
            using (var stream = System.IO.File.Create(file)) encoder.Save(stream);
        }
        finally { window.Close(); }
    }
}
