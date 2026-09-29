using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using ClashApi = Autodesk.Navisworks.Api.Clash;

namespace KPLN_NavisMcpBridge
{
    public sealed class TestRenameRequest
    {
        public string oldName { get; set; }
        public string newName { get; set; }
    }

    public sealed class TestRenameDto
    {
        public string id { get; set; }
        public string old_name { get; set; }
        public string new_name { get; set; }
        public string outcome { get; set; }
    }

    public sealed class TestRenameBatchDto
    {
        public bool dry_run { get; set; }
        public List<TestRenameDto> updates { get; set; } = new List<TestRenameDto>();
    }

    internal static partial class ClashService
    {
        public static TestRenameBatchDto RenameTests(List<TestRenameRequest> updates, bool dryRun)
        {
            if (updates == null || updates.Count == 0)
                throw new ArgumentException("updates must not be empty");
            var sources = new HashSet<string>(StringComparer.Ordinal);
            var destinations = new HashSet<string>(StringComparer.Ordinal);
            foreach (var update in updates)
            {
                if (update == null || string.IsNullOrWhiteSpace(update.oldName) ||
                    string.IsNullOrWhiteSpace(update.newName))
                    throw new ArgumentException("Each update must contain nonempty oldName and newName");
                if (!sources.Add(update.oldName) || !destinations.Add(update.newName))
                    throw new ArgumentException("Repeated source or destination name in updates");
            }

            var document = Application.ActiveDocument;
            if (document == null) throw new StateConflictException("No open Navisworks document");
            var testsData = ClashApi.DocumentClash.ClashInstance(document)?.TestsData;
            if (testsData == null) throw new StateConflictException("Clash Detective is unavailable");
            var tests = testsData.Tests.OfType<ClashApi.ClashTest>().ToList();
            var plan = new List<ClashApi.ClashTest>();
            var output = new TestRenameBatchDto { dry_run = dryRun };
            // Resolve the entire request before starting a transaction. Occupied names,
            // including swaps/chains, are rejected so no temporary names are necessary.
            foreach (var update in updates)
            {
                var matches = tests.Where(t => string.Equals(t.DisplayName, update.oldName,
                    StringComparison.Ordinal)).ToList();
                if (matches.Count == 0)
                    throw new ResourceNotFoundException($"Test not found: '{update.oldName}'");
                if (matches.Count != 1)
                    throw new StateConflictException($"Ambiguous test name: '{update.oldName}'");
                var test = matches[0];
                if (test.Guid == Guid.Empty)
                    throw new StateConflictException($"Test has no identity: '{update.oldName}'");
                if (tests.Any(t => t.Guid != test.Guid && string.Equals(t.DisplayName,
                    update.newName, StringComparison.Ordinal)))
                    throw new StateConflictException($"Destination name is occupied: '{update.newName}'");
                plan.Add(test);
                output.updates.Add(new TestRenameDto
                {
                    id = test.Guid.ToString("D"), old_name = update.oldName, new_name = update.newName,
                    outcome = update.oldName == update.newName ? "unchanged" : "planned"
                });
            }
            if (dryRun || output.updates.All(item => item.outcome == "unchanged")) return output;

            using (var transaction = document.BeginTransaction("KPLN: rename clash tests"))
            {
                for (int i = 0; i < plan.Count; i++)
                {
                    var item = output.updates[i];
                    if (item.outcome == "unchanged") continue;
                    testsData.TestsEditDisplayName(plan[i], item.new_name);
                }
                // Re-read by identity: the API may replace wrappers while editing.
                var actual = testsData.Tests.OfType<ClashApi.ClashTest>().ToDictionary(t => t.Guid);
                foreach (var item in output.updates)
                    if (!actual.TryGetValue(Guid.Parse(item.id), out var test) ||
                        !string.Equals(test.DisplayName, item.new_name, StringComparison.Ordinal))
                        throw new StateConflictException($"Rename verification failed: '{item.old_name}'");
                transaction.Commit();
            }
            foreach (var item in output.updates)
                if (item.outcome == "planned") item.outcome = "updated";
            return output;
        }
    }
}
