using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Threading;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalModel
{
    internal static class ScheduleCellPageReader
    {
        public const int MaxPageBytes = 200 * 1024;
        public const int MaxPageRows = 200;
        public const int TimeBudgetMilliseconds = 5000;

        // Read only the requested page. A row is committed atomically so next_offset never skips cells.
        public static JObject Read(
            JObject metadata,
            int firstRowIndex,
            int totalRows,
            IList<int> columnIndexes,
            int limit,
            int offset,
            Func<int, int, string> readCell,
            CancellationToken cancellationToken,
            int timeBudgetMilliseconds = TimeBudgetMilliseconds)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (limit <= 0 || offset < 0 || totalRows < 0)
                throw new ArgumentOutOfRangeException("Invalid schedule page bounds.");
            if (columnIndexes == null || readCell == null || metadata == null)
                throw new ArgumentNullException("Schedule page input is missing.");

            int pageLimit = Math.Min(limit, MaxPageRows);
            int rowsToRead = offset >= totalRows ? 0 : Math.Min(pageLimit, totalRows - offset);
            JObject result = (JObject)metadata.DeepClone();
            JArray rows = new JArray();
            result["items"] = rows;
            result["total_count"] = totalRows;
            result["limit"] = pageLimit;
            result["offset"] = offset;
            result["max_size_bytes"] = MaxPageBytes;
            result["cell_error_count"] = 0;
            result["stop_reason"] = "size_limit";
            SetProgress(result, totalRows, offset, 0);
            if (JsonByteCount(result) > MaxPageBytes)
                return Failure("schedule_columns_too_large",
                    "Column metadata exceeds 200 KB. Request fewer columnIndexes.", offset);

            Stopwatch stopwatch = Stopwatch.StartNew();
            string stopReason = null;
            int errorCount = 0;
            for (int rowOffset = 0; rowOffset < rowsToRead; rowOffset++)
            {
                int rowIndex = checked(firstRowIndex + offset + rowOffset);
                JArray cells = new JArray();
                int rowBytes = 0;
                int rowErrors = 0;
                foreach (int columnIndex in columnIndexes)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (stopwatch.ElapsedMilliseconds >= timeBudgetMilliseconds)
                    {
                        stopReason = "time_limit";
                        break;
                    }

                    JObject cell = new JObject { { "columnIndex", columnIndex } };
                    try
                    {
                        cell["text"] = readCell(rowIndex, columnIndex) ?? string.Empty;
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        cell["text"] = JValue.CreateNull();
                        cell["error"] = ex.Message;
                        rowErrors++;
                    }

                    int cellBytes = JsonByteCount(cell);
                    if (cellBytes > MaxPageBytes - rowBytes)
                    {
                        stopReason = "size_limit";
                        break;
                    }
                    rowBytes += cellBytes;
                    cells.Add(cell);
                }

                if (stopReason != null)
                    break;

                rows.Add(new JObject { { "rowIndex", rowIndex }, { "cells", cells } });
                result["cell_error_count"] = errorCount + rowErrors;
                SetProgress(result, totalRows, offset, rows.Count);
                // Reserve space for the longest stop_reason, including all JSON escaping and UTF-8 bytes.
                result["stop_reason"] = "size_limit";
                if (JsonByteCount(result) > MaxPageBytes)
                {
                    rows.RemoveAt(rows.Count - 1);
                    stopReason = "size_limit";
                    break;
                }
                errorCount += rowErrors;
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (rows.Count == 0 && rowsToRead > 0 && stopReason != null)
                return Failure(
                    stopReason == "size_limit" ? "schedule_row_too_large" : "schedule_row_timeout",
                    stopReason == "size_limit"
                        ? "One row exceeds the 200 KB page budget. Request fewer columnIndexes; no cells were skipped."
                        : "The row could not be read within the time budget. Request fewer columnIndexes; no rows were skipped.",
                    offset);

            SetProgress(result, totalRows, offset, rows.Count);
            result["cell_error_count"] = errorCount;
            result["stop_reason"] = stopReason
                ?? (offset + rows.Count < totalRows ? "row_limit" : "end");
            return result;
        }

        private static void SetProgress(JObject result, int totalRows, int offset, int count)
        {
            bool hasMore = offset < totalRows && count < totalRows - offset;
            result["returned_count"] = count;
            result["has_more"] = hasMore;
            result["next_offset"] = hasMore ? (JToken)(offset + count) : JValue.CreateNull();
        }

        private static JObject Failure(string code, string message, int offset)
        {
            return new JObject
            {
                { "error", message },
                { "error_code", code },
                { "offset", offset },
                { "max_size_bytes", MaxPageBytes }
            };
        }

        private static int JsonByteCount(JToken value)
        {
            return Encoding.UTF8.GetByteCount(JsonConvert.SerializeObject(value, Formatting.None));
        }
    }
}
