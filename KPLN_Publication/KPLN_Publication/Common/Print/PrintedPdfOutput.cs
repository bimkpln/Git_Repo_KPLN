using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace KPLN_Publication
{
    internal static class PrintedPdfOutput
    {
        internal static bool IsComplete(string path)
        {
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    if (stream.Length < 10)
                        return false;
                    var header = new byte[5];
                    stream.Read(header, 0, header.Length);
                    if (Encoding.ASCII.GetString(header) != "%PDF-")
                        return false;
                    var tail = new byte[(int)Math.Min(1024, stream.Length)];
                    stream.Seek(-tail.Length, SeekOrigin.End);
                    stream.Read(tail, 0, tail.Length);
                    return Encoding.ASCII.GetString(tail).TrimEnd().EndsWith("%%EOF", StringComparison.Ordinal);
                }
            }
            catch (IOException) { return false; }
            catch (UnauthorizedAccessException) { return false; }
        }

        internal static void WaitForFiles(IList<string> files, TimeSpan timeout)
        {
            var watch = Stopwatch.StartNew();
            while (true)
            {
                string[] pending = files.Where(path => !IsComplete(path)).ToArray();
                if (pending.Length == 0)
                    return;
                if (watch.Elapsed >= timeout)
                    throw new TimeoutException("Не завершена запись PDF:\n" + string.Join("\n", pending)
                        + "\nГотовые файлы сохранены. Проверьте очередь принтера перед повторной печатью.");
                Thread.Sleep(250);
            }
        }

        internal static void Combine(IList<string> files, string destination)
        {
            if (files.Count == 0)
                throw new ArgumentException("Нет PDF для объединения.", nameof(files));

            // При ошибке чтения любого листа сохраняем все исходники и прежний итоговый файл.
            string temporary = Path.Combine(Path.GetDirectoryName(destination), Path.GetRandomFileName() + ".pdf");
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew))
                using (var document = new iTextSharp.text.Document())
                using (var writer = new PdfCopy(document, stream))
                {
                    document.Open();
                    foreach (string file in files)
                    {
                        using (var reader = new PdfReader(file))
                            writer.AddDocument(reader);
                    }
                    document.Close();
                }

                if (File.Exists(destination))
                    File.Replace(temporary, destination, null);
                else
                    File.Move(temporary, destination);
            }
            finally
            {
                if (File.Exists(temporary))
                    File.Delete(temporary);
            }
        }
    }
}
