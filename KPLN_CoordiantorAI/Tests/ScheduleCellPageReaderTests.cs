using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using KPLN_CoordiantorAI.ExternalModel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

internal static class ScheduleCellPageReaderTests
{
    private static int Main()
    {
        Action[] tests =
        {
            ReadsOnlyRequestedRowsAndColumns,
            PagesPreserveUnicodeAndEscapedValues,
            ReportsCellErrorsWithoutLosingOtherCells,
            OversizedRowsNeverAdvanceTheOffset,
            RejectsOversizedMetadataBeforeReading,
            EmptyPagesDoNotReadCells,
            CancellationPropagates,
            TimeBudgetKeepsOnlyCompleteRows,
            ValidatesBounds
        };
        try
        {
            foreach (Action test in tests)
            {
                test();
                Console.WriteLine("PASS " + test.Method.Name);
            }
            Console.WriteLine("All " + tests.Length + " tests passed.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static JObject Read(int total, int limit, int offset, Func<int, int, string> cell,
        CancellationToken token = default(CancellationToken), int budget = 5000, JObject metadata = null)
    {
        return ScheduleCellPageReader.Read(metadata ?? new JObject(), 3, total, new[] { 2, 5 },
            limit, offset, cell, token, budget);
    }

    private static void ReadsOnlyRequestedRowsAndColumns()
    {
        var seen = new List<int>();
        int offset = 0, calls = 0, pages = 0;
        JObject metadata = new JObject { { "scheduleName", "Schedule" } };
        do
        {
            JObject page = Read(450, 500, offset, (row, col) =>
            {
                Check(row >= 3 + offset && row < 3 + Math.Min(offset + 200, 450), "Read outside page");
                Check(col == 2 || col == 5, "Wrong column index");
                calls++;
                return row + ":" + col;
            }, metadata: metadata);
            Check((int)page["limit"] == 200, "Limit not capped");
            foreach (JObject row in page["items"])
                seen.Add((int)row["rowIndex"]);
            Check((bool)page["has_more"] == (seen.Count < 450), "Wrong has_more");
            pages++;
            if (!(bool)page["has_more"]) break;
            offset = (int)page["next_offset"];
        } while (pages < 10);
        Check(pages == 3 && calls == 900, "Extra reads or wrong number of pages");
        Check(seen.SequenceEqual(Enumerable.Range(3, 450)), "Rows skipped or duplicated");
        Check(metadata["items"] == null, "Metadata mutated");
    }

    private static void PagesPreserveUnicodeAndEscapedValues()
    {
        string value = new string('\u042f', 3500) + "\"\\\r\n12,50 \u043c\u00b2";
        int offset = 0, pages = 0;
        var seen = new List<int>();
        while (pages++ < 100)
        {
            JObject page = Read(75, 200, offset, (r, c) => value);
            int bytes = Encoding.UTF8.GetByteCount(page.ToString(Formatting.None));
            Check(bytes <= ScheduleCellPageReader.MaxPageBytes, "Page exceeds byte limit");
            Check((int)page["returned_count"] > 0, "Pagination made no progress");
            foreach (JObject row in page["items"])
            {
                seen.Add((int)row["rowIndex"]);
                foreach (JObject cell in row["cells"])
                    Check((string)cell["text"] == value, "Text was truncated or changed");
            }
            if (!(bool)page["has_more"])
            {
                Check(page["next_offset"].Type == JTokenType.Null, "Final page has next offset");
                break;
            }
            Check((string)page["stop_reason"] == "size_limit", "Wrong stop reason");
            int next = (int)page["next_offset"];
            Check(next > offset, "Offset did not advance");
            offset = next;
        }
        Check(pages > 1 && seen.SequenceEqual(Enumerable.Range(3, 75)), "Byte pages lost rows");
    }

    private static void ReportsCellErrorsWithoutLosingOtherCells()
    {
        JObject page = Read(2, 200, 0, (row, col) =>
        {
            if (row == 3 && col == 5) throw new InvalidOperationException("Cell unavailable");
            return "2,50";
        });
        JToken broken = page["items"][0]["cells"][1];
        Check(broken["text"].Type == JTokenType.Null, "Error cell presented as empty value");
        Check((string)broken["error"] == "Cell unavailable", "Error missing");
        Check((int)page["cell_error_count"] == 1 && (int)page["returned_count"] == 2, "Lost rows");
        Check((string)page["items"][1]["cells"][1]["text"] == "2,50", "Other cell lost");
    }

    private static void OversizedRowsNeverAdvanceTheOffset()
    {
        Func<int, int, string> cell = (row, col) =>
            row == 4 && col == 5 ? new string('x', ScheduleCellPageReader.MaxPageBytes) : "ok";
        JObject first = Read(3, 200, 0, cell);
        Check((int)first["returned_count"] == 1 && (int)first["next_offset"] == 1, "Partial row committed");
        JObject second = Read(3, 200, 1, cell);
        Check((string)second["error_code"] == "schedule_row_too_large", "Oversize error missing");
        Check((int)second["offset"] == 1, "Failed row skipped");
        JObject narrowed = ScheduleCellPageReader.Read(new JObject(), 3, 3, new[] { 2 }, 200, 1,
            cell, CancellationToken.None);
        Check((int)narrowed["returned_count"] == 2, "Narrower columns cannot recover");
    }

    private static void RejectsOversizedMetadataBeforeReading()
    {
        JObject metadata = new JObject { { "columns", new string('x', ScheduleCellPageReader.MaxPageBytes) } };
        bool called = false;
        JObject page = Read(1, 200, 0, (r, c) => { called = true; return ""; }, metadata: metadata);
        Check(!called && (string)page["error_code"] == "schedule_columns_too_large", "Metadata budget ignored");
    }

    private static void EmptyPagesDoNotReadCells()
    {
        foreach (int total in new[] { 0, 10 })
        {
            JObject page = Read(total, 200, int.MaxValue, (r, c) => { throw new Exception("Unexpected read"); });
            Check((int)page["returned_count"] == 0 && !(bool)page["has_more"], "Empty page incorrect");
            Check(page["next_offset"].Type == JTokenType.Null, "Empty page has continuation");
        }
    }

    private static void CancellationPropagates()
    {
        using (var cts = new CancellationTokenSource())
        {
            int reads = 0;
            Expect<OperationCanceledException>(() => Read(10, 200, 0, (r, c) =>
            {
                reads++;
                cts.Cancel();
                return "value";
            }, cts.Token));
            Check(reads == 1, "Reads continued after cancellation");
            Expect<OperationCanceledException>(() => Read(10, 200, 0, (r, c) => "value", cts.Token));
        }
    }

    private static void TimeBudgetKeepsOnlyCompleteRows()
    {
        JObject none = Read(2, 200, 0, (r, c) => "value", budget: 0);
        Check((string)none["error_code"] == "schedule_row_timeout", "Timeout without progress not reported");
        JObject page = Read(3, 200, 0, (r, c) =>
        {
            if (r == 4 && c == 2) Thread.Sleep(250);
            return "value";
        }, budget: 200);
        Check((int)page["returned_count"] == 1 && (int)page["next_offset"] == 1, "Partial row committed on timeout");
        Check((string)page["stop_reason"] == "time_limit", "Timeout reason missing");
    }

    private static void ValidatesBounds()
    {
        Expect<ArgumentOutOfRangeException>(() => Read(1, 0, 0, (r, c) => ""));
        Expect<ArgumentOutOfRangeException>(() => Read(1, 200, -1, (r, c) => ""));
    }

    private static void Expect<T>(Action action) where T : Exception
    {
        try { action(); }
        catch (T) { return; }
        throw new Exception("Expected " + typeof(T).Name);
    }

    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }
}
