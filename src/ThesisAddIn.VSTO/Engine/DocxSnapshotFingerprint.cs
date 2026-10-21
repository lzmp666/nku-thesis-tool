using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ThesisAddIn.VSTO.Engine
{
    /// <summary>Stable DOCX fingerprint ignoring ZIP metadata and Word's volatile core timestamps.</summary>
    public static class DocxSnapshotFingerprint
    {
        public static string Compute(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("必须提供 DOCX 文件路径。", nameof(path));
            }

            using (var file = File.OpenRead(path))
            using (var archive = new ZipArchive(file, ZipArchiveMode.Read, leaveOpen: false))
            using (var fingerprint = IncrementalHash.CreateHash(HashAlgorithmName.SHA256))
            using (var entryHash = SHA256.Create())
            {
                foreach (ZipArchiveEntry entry in archive.Entries
                    .Where(item => !string.Equals(item.FullName, "docProps/core.xml", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(item => item.FullName, StringComparer.Ordinal))
                {
                    fingerprint.AppendData(Encoding.UTF8.GetBytes(entry.FullName));
                    fingerprint.AppendData(new byte[] { 0 });
                    using (Stream content = entry.Open())
                    {
                        fingerprint.AppendData(entryHash.ComputeHash(content));
                    }
                }

                return BitConverter.ToString(fingerprint.GetHashAndReset()).Replace("-", string.Empty).ToLowerInvariant();
            }
        }
    }
}

