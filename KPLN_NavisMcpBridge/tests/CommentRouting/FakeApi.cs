// Test doubles for the API surface used by the real comment service.
namespace Autodesk.Navisworks.Api
{
    public class Comment : IDisposable
    {
        public string Body { get; set; }
        public string Author { get; set; } = "test";
        public void Dispose() { }
    }
    public class CommentCollection : List<Comment>, IDisposable
    {
        public void CopyFrom(IEnumerable<Comment> items)
        {
            Clear();
            AddRange(items.Select(c => new Comment { Author = c.Author, Body = c.Body }));
        }
        public void Dispose() { }
    }
    public enum CommentStatus { New }
    public class SavedItem
    {
        public Guid Guid { get; set; } = Guid.NewGuid();
        public string DisplayName { get; set; }
        public SavedItem Parent { get; set; }
        public CommentCollection Comments { get; } = new CommentCollection();
    }
    public static class Application
    {
        public static Document ActiveDocument { get; set; }
    }
    public class Document
    {
        public List<SavedItem> Items { get; } = new();
        public int Transactions { get; private set; }
        public Transaction BeginTransaction(string name) { Transactions++; return new Transaction(this); }
        public Comment CreateCommentWithUniqueId(string text, CommentStatus status) => new() { Body = text };
    }
    public class Transaction : IDisposable
    {
        private readonly Action rollback;
        private bool committed;
        public Transaction(Document doc)
        {
            var states = doc.Items.Select(item => (
                item, comments: item.Comments.ToList(),
                status: (item as Clash.ClashResult)?.Status)).ToList();
            rollback = () =>
            {
                foreach (var state in states)
                {
                    state.item.Comments.CopyFrom(state.comments);
                    if (state.item is Clash.ClashResult result) result.Status = state.status.Value;
                }
            };
        }
        public void Commit() { committed = true; }
        public void Dispose() { if (!committed) rollback(); }
    }
}

namespace Autodesk.Navisworks.Api.Clash
{
    public interface IClashResult { }
    public enum ClashResultStatus { New, Active, Approved, Reviewed, Resolved }
    public class ClashResult : SavedItem, IClashResult
    {
        public ClashResultStatus Status { get; set; } = ClashResultStatus.New;
    }
    public class ClashResultGroup : SavedItem, IClashResult { }
    public class ClashTest : SavedItem { }
    public class DocumentClashTests
    {
        public List<SavedItem> Writes { get; } = new();
        public Guid FailOwner { get; set; }
        public void TestsEditResultComments(IClashResult target, CommentCollection comments)
        {
            var item = (SavedItem)target;
            if (item.Guid == FailOwner)
                throw new InvalidOperationException("write failed", new Exception("inner detail"));
            item.Comments.CopyFrom(comments);
            Writes.Add(item);
        }
        public void TestsEditResultStatus(IClashResult target, ClashResultStatus status)
        {
            ((ClashResult)target).Status = status;
        }
    }
}

namespace KPLN_NavisMcpBridge
{
    using Autodesk.Navisworks.Api;
    using Autodesk.Navisworks.Api.Clash;
    internal class ResourceNotFoundException : Exception
    {
        public ResourceNotFoundException(string message) : base(message) { }
    }
    internal class StateConflictException : Exception
    {
        public StateConflictException(string message) : base(message) { }
    }
    internal static partial class ClashService
    {
        internal static ClashTest Test;
        internal static DocumentClashTests Tests;
        private static DocumentClashTests FindManagedTest(string name, out ClashTest test)
        {
            if (name != Test.DisplayName) throw new ResourceNotFoundException(name);
            test = Test;
            return Tests;
        }
        private static void FindManagedResults(ClashTest test, HashSet<string> requested,
            Dictionary<string, List<ClashResult>> found)
        {
            foreach (var result in Application.ActiveDocument.Items.OfType<ClashResult>())
                if (requested.Contains(result.DisplayName)) found[result.DisplayName].Add(result);
        }
        private static ClashResultStatus ParseManagedStatus(string status) =>
            Enum.Parse<ClashResultStatus>(status, true);
    }
}
