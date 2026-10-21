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
            Assert.Empty(DocxSnapshotFingerprint.GetChangedParts(first, equivalent));
            Assert.Equal(new[] { "word/document.xml" }, DocxSnapshotFingerprint.GetChangedParts(first, changed));
            Assert.Contains("word/document.xml", DocxSnapshotFingerprint.DescribeDifferences(first, changed));
            Assert.DoesNotContain("正文", DocxSnapshotFingerprint.DescribeDifferences(first, changed));
        }
        finally
        {
            foreach (var path in new[] { first, equivalent, changed }) if (File.Exists(path)) File.Delete(path);
            Directory.Delete(directory);
        }
    }

    [Fact]
    public void IgnoresXmlAttributeOrderWhileDetectingTextWhitespaceAndFormattingChanges()
    {
        var directory = Path.Combine(Path.GetTempPath(), "nku-fingerprint-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var paths = new[] { "first", "reordered", "text", "format" }.Select(name => Path.Combine(directory, name + ".docx")).ToArray();
        try
        {
            WriteXmlPackage(paths[0], "<document a=\"1\" b=\"2\"><t xml:space=\"preserve\"> A </t></document>");
            WriteXmlPackage(paths[1], "<?xml version=\"1.0\"?><document b=\"2\" a=\"1\"><t xml:space=\"preserve\"> A </t></document>");
            WriteXmlPackage(paths[2], "<document a=\"1\" b=\"2\"><t xml:space=\"preserve\">A</t></document>");
            WriteXmlPackage(paths[3], "<document a=\"9\" b=\"2\"><t xml:space=\"preserve\"> A </t></document>");
            Assert.Equal(DocxSnapshotFingerprint.Compute(paths[0]), DocxSnapshotFingerprint.Compute(paths[1]));
            Assert.NotEqual(DocxSnapshotFingerprint.Compute(paths[0]), DocxSnapshotFingerprint.Compute(paths[2]));
            Assert.NotEqual(DocxSnapshotFingerprint.Compute(paths[0]), DocxSnapshotFingerprint.Compute(paths[3]));
        }
        finally
        {
            foreach (var path in paths) if (File.Exists(path)) File.Delete(path);
            Directory.Delete(directory);
        }
    }

    [Fact]
    public void PreservesMixedContentOrderAndNamespaceBindings()
    {
        var directory = Path.Combine(Path.GetTempPath(), "nku-fingerprint-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var paths = new[] { "first", "order", "binding" }.Select(name => Path.Combine(directory, name + ".docx")).ToArray();
        try
        {
            WriteXmlPackage(paths[0], "<document xmlns:p=\"urn:first\" type=\"p:X\">A<run>B</run>C</document>");
            WriteXmlPackage(paths[1], "<document xmlns:p=\"urn:first\" type=\"p:X\">AC<run>B</run></document>");
            WriteXmlPackage(paths[2], "<document xmlns:p=\"urn:second\" type=\"p:X\">A<run>B</run>C</document>");
            Assert.NotEqual(DocxSnapshotFingerprint.Compute(paths[0]), DocxSnapshotFingerprint.Compute(paths[1]));
            Assert.NotEqual(DocxSnapshotFingerprint.Compute(paths[0]), DocxSnapshotFingerprint.Compute(paths[2]));
        }
        finally
        {
            foreach (var path in paths) if (File.Exists(path)) File.Delete(path);
            Directory.Delete(directory);
        }
    }

    private static void WriteXmlPackage(string path, string xml)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        using var writer = new StreamWriter(archive.CreateEntry("word/document.xml").Open());
        writer.Write(xml);
    }

    [Fact]
    public void RegeneratedSessionIdsAndAutomaticSeparatorIdsAreNotContentChanges()
    {
        var directory = Path.Combine(Path.GetTempPath(), "nku-fingerprint-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var first = Path.Combine(directory, "first.docx");
        var second = Path.Combine(directory, "second.docx");
        try
        {
            WriteWordExportPackage(first, "00112233", "11223344");
            WriteWordExportPackage(second, "00998877", "77889900");
            Assert.Equal(DocxSnapshotFingerprint.Compute(first), DocxSnapshotFingerprint.Compute(second));
            Assert.Empty(DocxSnapshotFingerprint.GetChangedParts(first, second));
        }
        finally
        {
            foreach (var path in new[] { first, second }) if (File.Exists(path)) File.Delete(path);
            Directory.Delete(directory);
        }
    }

    [Fact]
    public void RetainsBodyIdentityFontValuesAndOrdinaryNoteContent()
    {
        var directory = Path.Combine(Path.GetTempPath(), "nku-fingerprint-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var paths = new[] { "first", "identity", "font", "note" }.Select(name => Path.Combine(directory, name + ".docx")).ToArray();
        try
        {
            WriteWordExportPackage(paths[0], "00112233", "11223344");
            WriteWordExportPackage(paths[1], "00998877", "77889900", bodyId: "55555555");
            WriteWordExportPackage(paths[2], "00998877", "77889900", font: "黑体");
            WriteWordExportPackage(paths[3], "00998877", "77889900", noteText: "修改后的普通脚注");
            foreach (var changed in paths.Skip(1)) Assert.NotEqual(DocxSnapshotFingerprint.Compute(paths[0]), DocxSnapshotFingerprint.Compute(changed));
        }
        finally
        {
            foreach (var path in paths) if (File.Exists(path)) File.Delete(path);
            Directory.Delete(directory);
        }
    }

    private static void WriteWordExportPackage(string path, string session, string separatorId,
        string bodyId = "12345678", string font = "宋体", string noteText = "普通脚注")
    {
        const string ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"";
        var parts = new[]
        {
            (Name: "word/document.xml", Xml: $"<w:document {ns}><w:body><w:p w:rsidRDefault=\"{session}\" w14:paraId=\"{bodyId}\"><w:r><w:rPr><w:rFonts w:eastAsia=\"{font}\"/></w:rPr><w:t>正文</w:t></w:r></w:p></w:body></w:document>"),
            (Name: "word/settings.xml", Xml: $"<w:settings {ns}><w:rsids><w:rsid w:val=\"{session}\"/></w:rsids><w:zoom w:percent=\"100\"/></w:settings>"),
            (Name: "word/footnotes.xml", Xml: Notes("footnote")),
            (Name: "word/endnotes.xml", Xml: Notes("endnote"))
        };
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var part in parts)
        {
            using var writer = new StreamWriter(archive.CreateEntry(part.Name).Open());
            writer.Write(part.Xml);
        }

        string Notes(string kind) => $"<w:{kind}s {ns}><w:{kind} w:type=\"separator\" w:id=\"-1\"><w:p w14:paraId=\"{separatorId}\" w:rsidR=\"{session}\" w:rsidRDefault=\"{session}\"><w:r><w:separator/></w:r></w:p></w:{kind}><w:{kind} w:id=\"1\"><w:p w14:paraId=\"33333333\"><w:r><w:t>{noteText}</w:t></w:r></w:p></w:{kind}></w:{kind}s>";
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

