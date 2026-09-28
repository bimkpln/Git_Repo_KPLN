using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using KPLN_NavisMcpBridge;

static class Program
{
    static void Assert(bool value, string message)
    {
        if (!value) throw new Exception(message);
    }
    static void Throws<T>(Action action) where T : Exception
    {
        try { action(); } catch (T) { return; }
        throw new Exception("Expected " + typeof(T).Name);
    }
    static void Reset()
    {
        Application.ActiveDocument = new Document();
        ClashService.Test = new ClashTest { DisplayName = "test" };
        ClashService.Tests = new DocumentClashTests();
        Application.ActiveDocument.Items.Add(ClashService.Test);
    }
    static T Add<T>(string name, SavedItem parent = null) where T : SavedItem, new()
    {
        var item = new T { DisplayName = name, Parent = parent ?? ClashService.Test };
        Application.ActiveDocument.Items.Add(item);
        return item;
    }
    static void Main()
    {
        Action[] tests = { MixedBatch, ExactDedupAndOldChild, SameNamesDifferentGroups,
            NearestGroup, ReadOnlyTargets, MissingAndAmbiguous, FailedOwner,
            SingleStatusAndGroupComment, SingleFailureRollsBack, EmptyComment };
        foreach (var test in tests) { Reset(); test(); Console.WriteLine("PASS " + test.Method.Name); }
        Console.WriteLine($"{tests.Length} comment routing tests passed (API doubles).");
    }
    static void MixedBatch()
    {
        var group = Add<ClashResultGroup>("group");
        var a = Add<ClashResult>("a", group);
        var b = Add<ClashResult>("b", group);
        var single = Add<ClashResult>("single");
        var sibling = Add<ClashResult>("not requested", group);
        sibling.Status = ClashResultStatus.Approved;
        var output = ClashService.AddResultComments("test", new[] { "a", "b", "a", "single" }, "note");
        Assert(output.targets.Count == 2 && output.updated.Count == 3, "one write per owner");
        Assert(group.Comments.Count == 1 && single.Comments.Count == 1, "owners receive comments");
        Assert(a.Comments.Count == 0 && b.Comments.Count == 0 && sibling.Comments.Count == 0, "no child writes");
        Assert(a.Status == ClashResultStatus.New && sibling.Status == ClashResultStatus.Approved, "preserve statuses");
        Assert(Application.ActiveDocument.Transactions == 1, "single transaction");
    }
    static void ExactDedupAndOldChild()
    {
        var group = Add<ClashResultGroup>("group");
        var a = Add<ClashResult>("a", group);
        a.Comments.Add(new Comment { Body = "note" });
        group.Comments.Add(new Comment { Body = "old" });
        ClashService.AddResultComments("test", new[] { "a" }, "note");
        var again = ClashService.AddResultComments("test", new[] { "a" }, "note");
        Assert(group.Comments.Select(c => c.Body).SequenceEqual(new[] { "old", "note" }), "preserve old, append exact");
        Assert(a.Comments.Count == 1 && again.unchanged.Count == 1, "old child retained, repeat skipped");
        ClashService.AddResultComments("test", new[] { "a" }, " note ");
        Assert(group.Comments.Last().Body == " note ", "do not trim body");
    }
    static void SameNamesDifferentGroups()
    {
        Add<ClashResult>("a", Add<ClashResultGroup>("same"));
        Add<ClashResult>("b", Add<ClashResultGroup>("same"));
        var output = ClashService.AddResultComments("test", new[] { "a", "b" }, "note");
        Assert(output.targets.Count == 2 && output.targets[0].id != output.targets[1].id, "use identity not name");
    }
    static void NearestGroup()
    {
        var outer = Add<ClashResultGroup>("outer");
        var inner = Add<ClashResultGroup>("inner", outer);
        Add<ClashResult>("a", inner);
        ClashService.AddResultComments("test", new[] { "a" }, "note");
        Assert(inner.Comments.Count == 1 && outer.Comments.Count == 0, "nearest group only");
    }
    static void ReadOnlyTargets()
    {
        var group = Add<ClashResultGroup>("group");
        Add<ClashResult>("a", group);
        group.Comments.Add(new Comment { Body = "old" });
        var output = ClashService.GetCommentTargets("test", new[] { "a" });
        Assert(output[0].kind == "group" && output[0].comments[0] == "test: old", "read owner comments");
        Assert(Application.ActiveDocument.Transactions == 0 && ClashService.Tests.Writes.Count == 0, "read only");
    }
    static void MissingAndAmbiguous()
    {
        var a = Add<ClashResult>("a");
        Throws<ResourceNotFoundException>(() => ClashService.AddResultComments("test", new[] { "a", "missing" }, "note"));
        Add<ClashResult>("a");
        Throws<StateConflictException>(() => ClashService.AddResultComments("test", new[] { "a" }, "note"));
        Assert(a.Comments.Count == 0 && Application.ActiveDocument.Transactions == 0, "resolve before writing");
    }
    static void FailedOwner()
    {
        var group = Add<ClashResultGroup>("group");
        Add<ClashResult>("a", group);
        Add<ClashResult>("b", group);
        var single = Add<ClashResult>("single");
        ClashService.Tests.FailOwner = group.Guid;
        var output = ClashService.AddResultComments("test", new[] { "a", "b", "single" }, "note");
        Assert(output.failed.Count == 2 && output.failed["a"].Contains("inner detail"), "full owner error for every child");
        Assert(single.Comments.Count == 1 && group.Comments.Count == 0, "other owner succeeds");
    }
    static void SingleStatusAndGroupComment()
    {
        var group = Add<ClashResultGroup>("group");
        var a = Add<ClashResult>("a", group);
        var b = Add<ClashResult>("b", group);
        ClashService.SetResultStatus("test", "a", "Active", "note");
        Assert(group.Comments.Count == 1 && a.Comments.Count == 0, "single endpoint routes comment");
        Assert(a.Status == ClashResultStatus.Active && b.Status == ClashResultStatus.New, "only requested status");
    }
    static void SingleFailureRollsBack()
    {
        var group = Add<ClashResultGroup>("group");
        var a = Add<ClashResult>("a", group);
        ClashService.Tests.FailOwner = group.Guid;
        Throws<InvalidOperationException>(() => ClashService.SetResultStatus("test", "a", "Active", "note"));
        Assert(a.Status == ClashResultStatus.New, "failed comment rolls back status");
    }
    static void EmptyComment()
    {
        Add<ClashResult>("a");
        Throws<ArgumentException>(() => ClashService.AddResultComments("test", new[] { "a" }, " "));
        Assert(Application.ActiveDocument.Transactions == 0, "empty comment rejected before write");
    }
}
