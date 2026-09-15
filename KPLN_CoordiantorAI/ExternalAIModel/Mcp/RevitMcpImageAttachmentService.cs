using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal sealed class RevitMcpImageAttachment
    {
        public string Id { get; set; }
        public string OwnerId { get; set; }
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string MimeType { get; set; }
        public long SizeBytes { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool IsTemporary { get; set; }
        public string PreparedDataUri { get; set; }
    }

    internal static class RevitMcpImageAttachmentService
    {
        public const long MaxImageBytes = 15L * 1024L * 1024L;
        public const int MaxImagesPerOwner = 5;
        public const int MaxImageDimension = 2048;
        private const long MaxPreparedBytes = 8L * 1024L * 1024L;

        private static readonly object SyncRoot = new object();
        private static readonly Dictionary<string, RevitMcpImageAttachment> Attachments =
            new Dictionary<string, RevitMcpImageAttachment>(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> SupportedExtensions =
            new HashSet<string>(new[] { ".png", ".jpg", ".jpeg", ".webp", ".bmp" }, StringComparer.OrdinalIgnoreCase);

        public static bool IsSupportedImagePath(string filePath)
        {
            return !string.IsNullOrWhiteSpace(filePath)
                && SupportedExtensions.Contains(Path.GetExtension(filePath) ?? string.Empty);
        }

        public static RevitMcpImageAttachment RegisterImageFile(string filePath, string ownerId)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("Image path is empty.", nameof(filePath));

            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath))
                throw new FileNotFoundException("The selected image no longer exists.", fullPath);
            if (!IsSupportedImagePath(fullPath))
                throw new NotSupportedException("Supported image formats: PNG, JPG, JPEG, WEBP and BMP.");

            return RegisterCore(fullPath, Path.GetFileName(fullPath), ownerId, false);
        }

        public static RevitMcpImageAttachment RegisterClipboardPng(byte[] pngBytes, string ownerId)
        {
            if (pngBytes == null || pngBytes.Length == 0)
                throw new ArgumentException("Clipboard image is empty.", nameof(pngBytes));
            if (pngBytes.LongLength > MaxImageBytes)
                throw new InvalidOperationException("The clipboard image is larger than the 15 MB limit.");

            string ownerFolder = GetOwnerTemporaryFolder(ownerId);
            Directory.CreateDirectory(ownerFolder);
            string fileName = "clipboard-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + ".png";
            string filePath = Path.Combine(ownerFolder, fileName);
            File.WriteAllBytes(filePath, pngBytes);

            try
            {
                return RegisterCore(filePath, fileName, ownerId, true);
            }
            catch
            {
                TryDeleteFile(filePath);
                throw;
            }
        }

        public static List<RevitMcpImageAttachment> GetImages(string ownerId)
        {
            lock (SyncRoot)
            {
                return Attachments.Values
                    .Where(item => string.Equals(item.OwnerId, ownerId, StringComparison.Ordinal))
                    .OrderBy(item => item.FileName, StringComparer.CurrentCultureIgnoreCase)
                    .Select(Clone)
                    .ToList();
            }
        }

        public static JArray CreateOpenAiImageContent(string ownerId)
        {
            JArray result = new JArray();
            foreach (RevitMcpImageAttachment image in GetImages(ownerId))
            {
                string dataUri = GetOrCreatePreparedDataUri(image.Id, ownerId);
                result.Add(new JObject
                {
                    { "type", "image_url" },
                    { "image_url", new JObject { { "url", dataUri } } }
                });
            }
            return result;
        }

        public static void RemoveImage(string ownerId, string imageId)
        {
            RevitMcpImageAttachment removed = null;
            lock (SyncRoot)
            {
                RevitMcpImageAttachment image;
                if (Attachments.TryGetValue(imageId ?? string.Empty, out image)
                    && string.Equals(image.OwnerId, ownerId, StringComparison.Ordinal))
                {
                    removed = image;
                    Attachments.Remove(imageId);
                }
            }

            if (removed != null && removed.IsTemporary)
                TryDeleteFile(removed.FilePath);
        }

        public static void RemoveOwnerImages(string ownerId)
        {
            foreach (RevitMcpImageAttachment image in GetImages(ownerId))
                RemoveImage(ownerId, image.Id);

            TryDeleteDirectory(GetOwnerTemporaryFolder(ownerId));
        }

        private static RevitMcpImageAttachment RegisterCore(
            string filePath,
            string fileName,
            string ownerId,
            bool isTemporary)
        {
            if (string.IsNullOrWhiteSpace(ownerId))
                throw new ArgumentException("Attachment owner id is empty.", nameof(ownerId));

            FileInfo info = new FileInfo(filePath);
            if (info.Length > MaxImageBytes)
                throw new InvalidOperationException("The image is larger than the 15 MB limit: " + fileName);

            PreparedImage prepared = PrepareImage(
                filePath,
                ResolveMimeType(Path.GetExtension(filePath)));
            string preparedDataUri = "data:"
                + prepared.MimeType
                + ";base64,"
                + Convert.ToBase64String(prepared.Bytes);
            lock (SyncRoot)
            {
                int ownerCount = Attachments.Values.Count(
                    item => string.Equals(item.OwnerId, ownerId, StringComparison.Ordinal));
                if (ownerCount >= MaxImagesPerOwner)
                    throw new InvalidOperationException("No more than 5 images can be attached to one question.");

                RevitMcpImageAttachment existing = Attachments.Values.FirstOrDefault(
                    item => string.Equals(item.OwnerId, ownerId, StringComparison.Ordinal)
                        && string.Equals(item.FilePath, filePath, StringComparison.OrdinalIgnoreCase));
                if (existing != null)
                    return Clone(existing);

                RevitMcpImageAttachment attachment = new RevitMcpImageAttachment
                {
                    Id = "image-" + Guid.NewGuid().ToString("N"),
                    OwnerId = ownerId,
                    FilePath = filePath,
                    FileName = fileName,
                    MimeType = prepared.MimeType,
                    SizeBytes = info.Length,
                    Width = prepared.Width,
                    Height = prepared.Height,
                    IsTemporary = isTemporary,
                    PreparedDataUri = preparedDataUri
                };
                Attachments[attachment.Id] = attachment;
                return Clone(attachment);
            }
        }

        private static string GetOrCreatePreparedDataUri(string imageId, string ownerId)
        {
            RevitMcpImageAttachment image;
            lock (SyncRoot)
            {
                if (!Attachments.TryGetValue(imageId, out image)
                    || !string.Equals(image.OwnerId, ownerId, StringComparison.Ordinal))
                    throw new FileNotFoundException("The attached image is no longer available.");
                if (!string.IsNullOrWhiteSpace(image.PreparedDataUri))
                    return image.PreparedDataUri;
                image = Clone(image);
            }

            PreparedImage prepared = PrepareImage(image.FilePath, image.MimeType);
            string dataUri = "data:" + prepared.MimeType + ";base64," + Convert.ToBase64String(prepared.Bytes);
            lock (SyncRoot)
            {
                RevitMcpImageAttachment current;
                if (Attachments.TryGetValue(imageId, out current)
                    && string.Equals(current.OwnerId, ownerId, StringComparison.Ordinal))
                {
                    current.PreparedDataUri = dataUri;
                    current.MimeType = prepared.MimeType;
                    current.Width = prepared.Width;
                    current.Height = prepared.Height;
                }
            }
            return dataUri;
        }

        private static PreparedImage PrepareImage(string filePath, string originalMimeType)
        {
            BitmapSource frame = ResizeToMaxDimension(
                LoadFirstFrame(filePath),
                MaxImageDimension);
            bool preferPng = string.Equals(originalMimeType, "image/png", StringComparison.OrdinalIgnoreCase);
            PreparedImage prepared = Encode(frame, preferPng);
            if (prepared.Bytes.LongLength > MaxPreparedBytes && preferPng)
                prepared = Encode(frame, false);
            if (prepared.Bytes.LongLength > MaxPreparedBytes)
                throw new InvalidOperationException("The prepared image is larger than the 8 MB request limit.");
            return prepared;
        }

        private static BitmapFrame LoadFirstFrame(string filePath)
        {
            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                BitmapImage image = new BitmapImage();
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.StreamSource = stream;
                image.EndInit();
                image.Freeze();
                return BitmapFrame.Create(image);
            }
        }

        private static BitmapSource ResizeToMaxDimension(BitmapSource source, int maxDimension)
        {
            int longestSide = Math.Max(source.PixelWidth, source.PixelHeight);
            if (longestSide <= maxDimension)
                return source;

            double scale = (double)maxDimension / longestSide;
            TransformedBitmap resized = new TransformedBitmap(
                source,
                new ScaleTransform(scale, scale));
            resized.Freeze();
            return resized;
        }

        private static PreparedImage Encode(BitmapSource source, bool asPng)
        {
            BitmapEncoder encoder = asPng
                ? (BitmapEncoder)new PngBitmapEncoder()
                : new JpegBitmapEncoder { QualityLevel = 85 };
            encoder.Frames.Add(BitmapFrame.Create(source));
            using (MemoryStream stream = new MemoryStream())
            {
                encoder.Save(stream);
                return new PreparedImage
                {
                    Bytes = stream.ToArray(),
                    MimeType = asPng ? "image/png" : "image/jpeg",
                    Width = source.PixelWidth,
                    Height = source.PixelHeight
                };
            }
        }

        private static string ResolveMimeType(string extension)
        {
            extension = (extension ?? string.Empty).ToLowerInvariant();
            return extension == ".jpg" || extension == ".jpeg" ? "image/jpeg" : "image/png";
        }

        private static string GetOwnerTemporaryFolder(string ownerId)
        {
            string safeOwnerId = new string((ownerId ?? string.Empty).Where(char.IsLetterOrDigit).ToArray());
            if (string.IsNullOrWhiteSpace(safeOwnerId))
                safeOwnerId = "unknown";
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "KPLN", "CoordinatorAI", "Attachments", safeOwnerId);
        }

        private static void TryDeleteFile(string filePath)
        {
            try { if (File.Exists(filePath)) File.Delete(filePath); } catch { }
        }

        private static void TryDeleteDirectory(string directoryPath)
        {
            try { if (Directory.Exists(directoryPath)) Directory.Delete(directoryPath, false); } catch { }
        }

        private static RevitMcpImageAttachment Clone(RevitMcpImageAttachment source)
        {
            return new RevitMcpImageAttachment
            {
                Id = source.Id,
                OwnerId = source.OwnerId,
                FilePath = source.FilePath,
                FileName = source.FileName,
                MimeType = source.MimeType,
                SizeBytes = source.SizeBytes,
                Width = source.Width,
                Height = source.Height,
                IsTemporary = source.IsTemporary,
                PreparedDataUri = source.PreparedDataUri
            };
        }

        private sealed class PreparedImage
        {
            public byte[] Bytes { get; set; }
            public string MimeType { get; set; }
            public int Width { get; set; }
            public int Height { get; set; }
        }
    }
}
