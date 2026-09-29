using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using KPLN_NavisMcpBridge;

static class Program
{
    static void Assert(bool ok, string message) { if (!ok) throw new Exception(message); }
    static void Throws<T>(Action action) where T : Exception
    { try { action(); } catch (T) { return; } throw new Exception("Expected " + typeof(T).Name); }
    static List<TestRenameRequest> Request(params string[] names) => Enumerable.Range(0, names.Length / 2)
        .Select(i => new TestRenameRequest { oldName = names[2 * i], newName = names[2 * i + 1] }).ToList();
    static DocumentClashTests Data => Application.ActiveDocument.Clash.TestsData;
    static ClashTest Add(string name)
    { var t = new ClashTest { DisplayName = name }; Data.Tests.Add(t); return t; }
    static void Main()
    {
        Action[] checks = { Preview, Apply, Missing, Ambiguous, Occupied, DuplicateRequest,
            Invalid, Unchanged, Swap, RollbackOnWriteFailure, RollbackOnVerificationFailure, Retry };
        foreach (var check in checks)
        {
            Application.ActiveDocument = new Document();
            check(); Console.WriteLine("PASS " + check.Method.Name);
        }
        Console.WriteLine($"{checks.Length} rename checks passed (API doubles).");
    }
    static void Preview()
    {
        var t = Add("АР / #1");
        var r = ClashService.RenameTests(Request(t.DisplayName, "АРХИВ_АР / #1"), true);
        Assert(r.dry_run && r.updates[0].outcome == "planned", "preview response");
        Assert(t.DisplayName == "АР / #1" && Application.ActiveDocument.Transactions == 0, "preview does not write");
    }
    static void Apply()
    {
        var a = Add("АР"); var b = Add(" ВК "); var numeric = Add("1.КР");
        var id = a.Guid; var results = a.Results;
        var r = ClashService.RenameTests(Request("АР", "АРХИВ_АР", " ВК ", "АРХИВ_ ВК "), false);
        Assert(a.DisplayName == "АРХИВ_АР" && b.DisplayName == "АРХИВ_ ВК ", "exact unicode names");
        Assert(a.Guid == id && ReferenceEquals(a.Results, results) && a.Status == "OK" && numeric.DisplayName == "1.КР", "preserve identity/results/status/other test");
        Assert(r.updates.All(x => x.outcome == "updated") && Application.ActiveDocument.Transactions == 1, "one transaction");
    }
    static void Missing()
    {
        var a = Add("a");
        Throws<ResourceNotFoundException>(() => ClashService.RenameTests(Request("a", "x", "missing", "y"), false));
        Assert(a.DisplayName == "a" && Data.Writes == 0, "late missing prevents first write");
    }
    static void Ambiguous()
    {
        Add("a"); Add("a");
        Throws<StateConflictException>(() => ClashService.RenameTests(Request("a", "x"), false));
        Assert(Data.Writes == 0, "duplicate existing names rejected");
    }
    static void Occupied()
    {
        Add("a"); Add("b"); Add("taken");
        Throws<StateConflictException>(() => ClashService.RenameTests(Request("a", "x", "b", "taken"), false));
        Assert(Data.Writes == 0, "late occupied name prevents first write");
    }
    static void DuplicateRequest()
    {
        Add("a"); Add("b");
        Throws<ArgumentException>(() => ClashService.RenameTests(Request("a", "x", "b", "x"), false));
        Throws<ArgumentException>(() => ClashService.RenameTests(Request("a", "x", "a", "y"), false));
        Assert(Data.Writes == 0, "repeated input rejected");
    }
    static void Invalid()
    {
        Throws<ArgumentException>(() => ClashService.RenameTests(null, false));
        Throws<ArgumentException>(() => ClashService.RenameTests(new(), false));
        Throws<ArgumentException>(() => ClashService.RenameTests(new() { null }, false));
        Throws<ArgumentException>(() => ClashService.RenameTests(Request("a", " \n"), false));
        Application.ActiveDocument = null;
        Throws<StateConflictException>(() => ClashService.RenameTests(Request("a", "b"), false));
    }
    static void Unchanged()
    {
        Add("a");
        var r = ClashService.RenameTests(Request("a", "a"), false);
        Assert(r.updates[0].outcome == "unchanged" && Application.ActiveDocument.Transactions == 0, "unchanged no write");
    }
    static void Swap()
    {
        Add("a"); Add("b");
        Throws<StateConflictException>(() => ClashService.RenameTests(Request("a", "b", "b", "a"), false));
        Assert(Data.Writes == 0, "swap rejected");
    }
    static void RollbackOnWriteFailure()
    {
        var a = Add("a"); var b = Add("b"); Data.FailAt = 2;
        Throws<InvalidOperationException>(() => ClashService.RenameTests(Request("a", "x", "b", "y"), false));
        Assert(a.DisplayName == "a" && b.DisplayName == "b" && Application.ActiveDocument.Commits == 0, "all writes rolled back");
    }
    static void RollbackOnVerificationFailure()
    {
        var a = Add("a"); Data.IgnoreWrites = true;
        Throws<StateConflictException>(() => ClashService.RenameTests(Request("a", "x"), false));
        Assert(a.DisplayName == "a" && Application.ActiveDocument.Commits == 0, "verification required");
    }
    static void Retry()
    {
        Add("a"); var request = Request("a", "АРХИВ_a");
        ClashService.RenameTests(request, false);
        Throws<ResourceNotFoundException>(() => ClashService.RenameTests(request, false));
        Assert(Data.Tests[0].DisplayName == "АРХИВ_a" && Data.Writes == 1, "stale retry cannot add prefix twice");
    }
}

namespace KPLN_NavisMcpBridge
{
    public class ResourceNotFoundException : Exception { public ResourceNotFoundException(string m) : base(m) {} }
    public class StateConflictException : Exception { public StateConflictException(string m) : base(m) {} }
}
namespace Autodesk.Navisworks.Api
{
    public class SavedItem { public Guid Guid { get; set; } = Guid.NewGuid(); public string DisplayName { get; set; } }
    public static class Application { public static Document ActiveDocument { get; set; } }
    public class Document
    {
        public Clash.DocumentClash Clash = new();
        public int Transactions, Commits;
        public Transaction BeginTransaction(string name) { Transactions++; return new Transaction(this); }
    }
    public class Transaction : IDisposable
    {
        readonly Document doc;
        readonly Dictionary<SavedItem, string> names;
        bool committed;
        public Transaction(Document d) { doc = d; names = d.Clash.TestsData.Tests.ToDictionary(t => t, t => t.DisplayName); }
        public void Commit() { committed = true; doc.Commits++; }
        public void Dispose() { if (!committed) foreach (var pair in names) pair.Key.DisplayName = pair.Value; }
    }
}
namespace Autodesk.Navisworks.Api.Clash
{
    public class ClashTest : SavedItem { public object Results = new(); public string Status = "OK"; }
    public class DocumentClash
    {
        public DocumentClashTests TestsData = new();
        public static DocumentClash ClashInstance(Document d) => d.Clash;
    }
    public class DocumentClashTests
    {
        public List<SavedItem> Tests = new();
        public int Writes, FailAt;
        public bool IgnoreWrites;
        public void TestsEditDisplayName(SavedItem test, string name)
        { Writes++; if (Writes == FailAt) throw new InvalidOperationException("Injected write failure"); if (!IgnoreWrites) test.DisplayName = name; }
    }
}
