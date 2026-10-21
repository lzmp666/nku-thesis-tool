using System.IO.Compression;
using ThesisAddIn.VSTO.Engine;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class DocxSnapshotFingerprintTests
{
    [Fact]
    public void FingerprintIgnoresZipMetadataButDetectsDocumentChanges()
    {
        var directory = Path.Combine(Path.GetTempPath(), "nku-fingerprint-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var first = Path.Combine(directory, "first.docx");
        var equivalent = Path.Combine(directory, "equivalent.docx");
        var changed = Path.Combine(directory, "changed.docx");

        try
        {
            WritePackage(first, "正文 A", "2026-10-20T01:00:00Z", reverseEntries: false);
            WritePackage(equivalent, "正文 A", "2026-10-21T01:00:00Z", reverseEntries: true);
            WritePackage(changed, "正文 B", "2026-10-20T01:00:00Z", reverseEntries: false);

            Assert.Equal(DocxSnapshotFingerprint.Compute(first), DocxSnapshotFingerprint.Compute(equivalent));
            Assert.NotEqual(DocxSnapshotFingerprint.Compute(first), DocxSnapshotFingerprint.Compute(changed));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static void WritePackage(string path, string documentText, string modified, bool reverseEntries)
    {
        var entries = new[]
        {
            (Name: "word/document.xml", Content: "<document>" + documentText + "</document>"),
            (Name: "docProps/core.xml", Content: "<core><modified>" + modified + "</modified></core>"),
            (Name: "word/comments.xml", Content: "<comments>preserve</comments>")
        };

        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var item in reverseEntries ? entries.Reverse() : entries)
        {
            var entry = archive.CreateEntry(item.Name);
            using var writer = new StreamWriter(entry.Open());
            writer.Write(item.Content);
        }
    }
}

