using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using ClashApi = Autodesk.Navisworks.Api.Clash;

namespace KPLN_NavisMcpBridge
{
    public class CommentTargetDto
    {
        public string kind { get; set; }
        public string name { get; set; }
        public string id { get; set; }
        public List<string> result_names { get; set; }
        public List<string> comments { get; set; }
        public string outcome { get; set; }
        public string error { get; set; }
    }

    public class ResultCommentBatchDto
    {
        // These lists retain requested result names for existing clients; targets lists actual owners.
        public List<string> updated { get; set; } = new List<string>();
        public List<string> unchanged { get; set; } = new List<string>();
        public Dictionary<string, string> failed { get; set; } =
            new Dictionary<string, string>(StringComparer.Ordinal);
        public List<CommentTargetDto> targets { get; set; } = new List<CommentTargetDto>();
    }

    internal static partial class ClashService
    {
        private sealed class CommentTarget
        {
            public SavedItem Owner { get; }
            public List<string> Results { get; } = new List<string>();

            public CommentTarget(SavedItem owner) { Owner = owner; }
        }

        private static SavedItem CommentOwner(ClashApi.ClashResult result)
        {
            for (var parent = result.Parent; parent != null; parent = parent.Parent)
                if (parent is ClashApi.ClashResultGroup) return parent;
            return result;
        }

        private static List<CommentTarget> BuildCommentTargets(IEnumerable<ClashApi.ClashResult> results)
        {
            var targets = new Dictionary<Guid, CommentTarget>();
            foreach (var result in results)
            {
                var owner = CommentOwner(result);
                var id = owner.Guid;
                if (id == Guid.Empty)
                    throw new StateConflictException($"Comment owner '{owner.DisplayName}' has no identity");
                if (!targets.TryGetValue(id, out var target))
                {
                    target = new CommentTarget(owner);
                    targets.Add(id, target);
                }
                target.Results.Add(result.DisplayName);
            }
            return targets.Values.ToList();
        }

        private static CommentTargetDto DescribeCommentTarget(CommentTarget target)
        {
            return new CommentTargetDto
            {
                kind = target.Owner is ClashApi.ClashResultGroup ? "group" : "result",
                name = target.Owner.DisplayName,
                id = target.Owner.Guid.ToString("D"),
                result_names = target.Results.ToList(),
                comments = target.Owner.Comments.Select(c => $"{c.Author}: {c.Body}").ToList()
            };
        }

        public static List<CommentTargetDto> GetCommentTargets(string testName, IEnumerable<string> resultNames)
        {
            var results = ResolveCommentResults(testName, resultNames, out var testsData);
            return BuildCommentTargets(results).Select(DescribeCommentTarget).ToList();
        }

        public static void SetResultStatus(string testName, string resultName, string status, string comment)
        {
            var parsed = string.IsNullOrEmpty(status)
                ? (ClashApi.ClashResultStatus?)null : ParseManagedStatus(status);
            if (!string.IsNullOrEmpty(comment) && string.IsNullOrWhiteSpace(comment))
                throw new ArgumentException("Comment must not contain only whitespace");
            var results = ResolveCommentResults(testName, new[] { resultName }, out var testsData);
            var result = results[0];
            var owner = CommentOwner(result);
            using (var transaction = Application.ActiveDocument.BeginTransaction("KPLN: edit clash result"))
            {
                // Status is always applied only to the requested result, never to its group.
                if (parsed.HasValue && result.Status != parsed.Value)
                    testsData.TestsEditResultStatus((ClashApi.IClashResult)result, parsed.Value);
                if (!string.IsNullOrEmpty(comment)) AppendComment(testsData, owner, comment);
                transaction.Commit();
            }
        }

        private static List<ClashApi.ClashResult> ResolveCommentResults(
            string testName, IEnumerable<string> resultNames, out ClashApi.DocumentClashTests testsData)
        {
            if (string.IsNullOrWhiteSpace(testName)) throw new ArgumentException("testName is required");
            if (resultNames == null) throw new ArgumentException("resultNames is required");
            var requested = new HashSet<string>(StringComparer.Ordinal);
            foreach (var name in resultNames)
            {
                if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Empty result name");
                requested.Add(name);
            }
            if (requested.Count == 0) throw new ArgumentException("resultNames must not be empty");
            testsData = FindManagedTest(testName, out var test);
            var found = requested.ToDictionary(name => name,
                name => new List<ClashApi.ClashResult>(), StringComparer.Ordinal);
            FindManagedResults(test, requested, found);
            var missing = found.Where(pair => pair.Value.Count == 0).Select(pair => pair.Key).ToList();
            if (missing.Count > 0)
                throw new ResourceNotFoundException($"Results not found in '{testName}': {string.Join(", ", missing)}");
            var ambiguous = found.Where(pair => pair.Value.Count > 1).Select(pair => pair.Key).ToList();
            if (ambiguous.Count > 0)
                throw new StateConflictException($"Duplicate result names in '{testName}': {string.Join(", ", ambiguous)}");
            return found.Values.Select(items => items[0]).ToList();
        }

        private static bool AppendComment(
            ClashApi.DocumentClashTests testsData, SavedItem owner, string text)
        {
            // Check only the owner: old child comments must not prevent adding a group comment.
            if (owner.Comments.Any(item => string.Equals(item.Body, text, StringComparison.Ordinal)))
                return false;
            using (var comments = new CommentCollection())
            using (var comment = Application.ActiveDocument.CreateCommentWithUniqueId(text, CommentStatus.New))
            {
                comments.CopyFrom(owner.Comments);
                comments.Add(comment);
                testsData.TestsEditResultComments((ClashApi.IClashResult)owner, comments);
            }
            return true;
        }

        public static ResultCommentBatchDto AddResultComments(
            string testName, IEnumerable<string> resultNames, string comment)
        {
            if (string.IsNullOrWhiteSpace(comment)) throw new ArgumentException("comment is required");
            var results = ResolveCommentResults(testName, resultNames, out var testsData);
            var targets = BuildCommentTargets(results);
            var output = new ResultCommentBatchDto();
            using (var transaction = Application.ActiveDocument.BeginTransaction("KPLN: add clash comments"))
            {
                foreach (var target in targets)
                {
                    var dto = DescribeCommentTarget(target);
                    output.targets.Add(dto);
                    try
                    {
                        var added = AppendComment(testsData, target.Owner, comment);
                        dto.outcome = added ? "updated" : "unchanged";
                        dto.comments = target.Owner.Comments.Select(c => $"{c.Author}: {c.Body}").ToList();
                        if (added) output.updated.AddRange(target.Results);
                        else output.unchanged.AddRange(target.Results);
                    }
                    catch (Exception ex)
                    {
                        dto.outcome = "failed";
                        dto.error = ex.ToString();
                        foreach (var name in target.Results) output.failed[name] = dto.error;
                    }
                }
                transaction.Commit();
            }
            return output;
        }
    }
}
