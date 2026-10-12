using System;
using System.Collections.Generic;
using System.Linq;
using KPLN_Tools_OVVK.Forms.Entities;
using KPLN_Tools_OVVK.Forms.Entities.VentConfigurator;
using KPLN_Tools_OVVK.Services.VentConfigurator;

internal static class DomainTests
{
    private static int _checks;

    private static void Check(bool condition, string name)
    {
        if (!condition)
            throw new Exception(name);

        _checks++;
    }

    private static void Reject(Action action, string name)
    {
        try
        {
            action();
        }
        catch (ArgumentException)
        {
            _checks++;
            return;
        }
        catch (InvalidOperationException)
        {
            _checks++;
            return;
        }

        throw new Exception(name);
    }

    internal static SectionCatalog Catalog()
    {
        var catalog = new SectionCatalog
        {
            IntermediateCount = 2,
            SourceTypeName = VentConfiguratorSettings.BaseFamilyTypeName,
            InstallationWidthMm = 1000,
            InstallationHeightMm = 1200,
            FrameHeightMm = 100
        };
        var choice = new SectionTypeChoice("Test", "Section");
        foreach (int slot in Enumerable.Range(1, 12))
        {
            catalog.Add(slot, new SectionDefinition
            {
                ParameterName = InstallationConfiguration.ParameterName(slot),
                DisplayName = "Section " + slot,
                Choices = new List<SectionTypeChoice> { choice },
                SourceType = choice,
                DefaultLengthMm = 200,
                DefaultWidthMm = 1000,
                DefaultHeightMm = 1200,
                DefaultOffsetXMm = 0,
                DefaultOffsetYMm = 0
            });
        }

        foreach (string name in SharedParameters.LengthNames)
            catalog.SharedLengthsMm[name] = 100;

        foreach (string name in SharedParameters.BooleanNames)
            catalog.SharedFlags[name] = false;

        return catalog;
    }

    private static void Main()
    {
        var configuration = InstallationConfiguration.CreateDraft(Catalog());
        configuration.Validate();
        Check(configuration.Blocks.Count == 4, "Initial composition");
        var first = configuration.Blocks[configuration.FirstSectionIndex];
        configuration.Move(configuration.FirstSectionIndex, configuration.FirstSectionIndex + 1);
        Check(ReferenceEquals(first, configuration.Blocks[configuration.FirstSectionIndex + 1]), "Reordering preserves values");
        configuration.Move(0, 2);
        Check(configuration.Blocks[0].IsConnector, "Connector cannot move");

        configuration.Resize(configuration.Capacity);
        Check(configuration.SectionCount == configuration.Capacity, "Resize to capacity");
        Reject(() => configuration.Resize(configuration.Capacity + 1), "Capacity overflow");
        configuration.Resize(configuration.Minimum);
        configuration.SetValve(true);
        Check(configuration.HasValve && configuration.Blocks[1].IsValve, "Valve position");
        configuration.Blocks[1].Length.Text = "357";
        configuration.SetValve(false);
        configuration.SetValve(true);
        Check(configuration.Blocks[1].Length.Text == "357", "Valve state survives toggling");

        var copy = configuration.Copy();
        Check(copy.Signature() == configuration.Signature(), "Deep copy signature");
        copy.Blocks[1].Length.Text = "600";
        copy.Info.Mark = "Copied";
        Check(configuration.Blocks[1].Length.Text == "357" && configuration.Info.Mark != "Copied", "Deep copy independence");

        var item = new FamilyTypeItem { Name = "Test", Configuration = configuration, PersistedName = "Test" };
        item.MarkSaved();
        Check(!item.IsDirty && !item.HasUserChanges, "Saved clean state");
        configuration.Info.Mark = "Edited";
        item.MarkChanged();
        Check(item.IsDirty && item.HasUserChanges, "Dirty state");
        item.MarkImported();
        item.MarkChanged();
        Check(item.IsDirty && !item.HasUserChanges, "Imported type still needs saving");
        var typeCopy = item.CreateCopy(new[] { item });
        Check(typeCopy.PersistedName == null && typeCopy.Configuration.CopySourceTypeName == "Test", "Copy creates a new type");
        Check(FamilyTypeItem.AvailableAutomaticName("Test", new[] { "test", "Test 1" }) == "Test 2", "Unique automatic names");

        string basis = VentConfiguratorSettings.BaseFamilyTypeName;
        var names = new[] { basis, "A", "B" };
        Check(TypeWritePlan.Make(names, "C", null, false).Create, "Create plan");
        Check(!TypeWritePlan.Make(names, "C", "A", false).Create, "Rename plan");
        Reject(() => TypeWritePlan.Make(names, basis, "A", false), "Protect template");
        Reject(() => TypeWritePlan.Make(names, "B", "A", false), "Protect other types");
        Reject(() => TypeWritePlan.Make(names, "C", "Missing", false), "Missing persisted type");

        foreach (var kind in new[] { RequestKind.OpenInRevit, RequestKind.AddToProject })
        {
            var plan = FamilyActionPlan.Make(kind, "saved.rfa", "project.rfa", "A", false);
            Check(!plan.WritesFamily && plan.Path == "saved.rfa", kind + " cannot save");
            Reject(() => FamilyActionPlan.Make(kind, "saved.rfa", null, "A", true), kind + " rejects dirty types");
        }

        var save = FamilyActionPlan.Make(RequestKind.Save, "saved.rfa", "project.rfa", "A", false);
        Check(save.WritesFamily && save.Path == "project.rfa", "Project save destination");
        Check(FamilyActionPlan.Make(RequestKind.Save, null, null, null, true).NeedsPath, "New file path required");
        Reject(() => VentTypeName.ValidateTypeName(basis), "Reject template name");
        Reject(() => VentTypeName.ValidateTypeName(" "), "Reject blank type name");

        var project = new ProjectItem { Name = "A", Code = "A", Paths = new[] { @"C:\Projects\A" } };
        Check(ProjectPathMatcher.Find(new[] { project }, @"C:\Projects\A\model.rvt").Project == project, "Project folder match");
        Check(ProjectPathMatcher.Find(new[] { project }, @"C:\Projects\AB\model.rvt").Project == null, "Folder boundary match");
        var duplicate = new ProjectItem { Name = "B", Code = "B", Paths = project.Paths };
        Check(ProjectPathMatcher.Find(new[] { project, duplicate }, @"C:\Projects\A\model.rvt").IsAmbiguous, "Ambiguous project match");

        var model = new OV_VentConfiguratorM();
        int notifications = 0;
        model.PropertyChanged += (sender, args) => notifications++;
        model.Configuration = configuration;
        model.IsBusy = true;
        Check(!model.CanEdit && !model.HasConfiguration, "Busy disables editing");
        model.IsBusy = false;
        Check(model.CanEdit && model.HasConfiguration && notifications == 2, "State change notifications");
        Console.WriteLine("PASS: " + _checks + " domain checks.");
    }
}
