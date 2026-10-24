using System.IO.Compression;
using ThesisAddIn.VSTO.Engine;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class DocxSnapshotFingerprintTests
{
    [Fact]
    public void OwnedCommentRunReserializationDoesNotBlockRepairsButTextAndHumanCommentsAreProtected()
    {
        const string ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        var directory = Path.Combine(Path.GetTempPath(), "nku-comment-fingerprint-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var paths = new[] { "first", "reserialized", "edited", "human" }.Select(name => Path.Combine(directory, name + ".docx")).ToArray();
        try
        {
            for (int i = 0; i < paths.Length; i++)
            {
                string body = i == 0
                    ? "<w:r><w:t>期望值：宋体</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>[NKU-TOOL-ANNOTATION:v1]</w:t></w:r>"
                    : "<w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/></w:rPr><w:t>期望值：</w:t></w:r><w:r><w:t>" +
                        (i == 2 ? "黑体" : "宋体") + "[NKU-TOOL-ANNOTATION:v1]</w:t></w:r>";
                using var archive = ZipFile.Open(paths[i], ZipArchiveMode.Create);
                using var writer = new StreamWriter(archive.CreateEntry("word/comments.xml").Open());
                writer.Write($"<w:comments {ns}><w:comment w:id=\"1\" w:author=\"NKU Thesis Tool\"><w:p>{body}</w:p></w:comment>" +
                    $"<w:comment w:id=\"2\" w:author=\"Reviewer\"><w:p><w:r><w:rPr><w:b w:val=\"{(i == 3 ? "1" : "0")}\"/></w:rPr><w:t>人工批注</w:t></w:r></w:p></w:comment></w:comments>");
            }
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
            Assert.True(DocxSnapshotFingerprint.Compute(paths[0]) == DocxSnapshotFingerprint.Compute(paths[1]),
                DocxSnapshotFingerprint.DescribeDifferences(paths[0], paths[1]));
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
    public void IgnoresWordProofingMetadataAndXmlIndentationButDetectsAuthoredTextChanges()
    {
        const string ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        var directory = Path.Combine(Path.GetTempPath(), "nku-fingerprint-proof-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var paths = new[] { "first", "word-export", "text-change" }.Select(name => Path.Combine(directory, name + ".docx")).ToArray();
        try
        {
            WriteWordParts(paths[0],
                $"<w:document {ns}><w:body><w:p><w:r><w:t>论文正文</w:t></w:r></w:p></w:body></w:document>",
                $"<w:settings {ns}><w:zoom w:percent=\"100\"/></w:settings>");
            WriteWordParts(paths[1],
                $"<w:document {ns}>\n  <w:body>\n    <w:p><w:pPr><w:rPr><w:lang w:eastAsia=\"zh-CN\"/></w:rPr></w:pPr><w:bookmarkStart w:id=\"1\" w:name=\"_NKU_MARK_test\"/><w:proofErr w:type=\"spellStart\"/><w:r><w:rPr><w:lang w:val=\"en-US\"/></w:rPr><w:t>论文</w:t></w:r><w:commentRangeStart w:id=\"2\"/><w:r><w:rPr><w:highlight w:val=\"yellow\"/></w:rPr><w:t>正文</w:t></w:r><w:commentRangeEnd w:id=\"2\"/><w:bookmarkEnd w:id=\"1\"/><w:proofErr w:type=\"spellEnd\"/></w:p>\n  </w:body>\n</w:document>",
                $"<w:settings {ns}><w:proofState w:spelling=\"clean\" w:grammar=\"dirty\"/><w:zoom w:percent=\"100\"/></w:settings>");
            WriteWordParts(paths[2],
                $"<w:document {ns}><w:body><w:p><w:proofErr w:type=\"spellStart\"/><w:r><w:t>论文正文已修改</w:t></w:r><w:proofErr w:type=\"spellEnd\"/></w:p></w:body></w:document>",
                $"<w:settings {ns}><w:proofState w:spelling=\"clean\"/><w:zoom w:percent=\"100\"/></w:settings>");

            Assert.True(DocxSnapshotFingerprint.Compute(paths[0]) == DocxSnapshotFingerprint.Compute(paths[1]),
                DocxSnapshotFingerprint.DescribeDifferences(paths[0], paths[1]));
            Assert.NotEqual(DocxSnapshotFingerprint.Compute(paths[0]), DocxSnapshotFingerprint.Compute(paths[2]));
        }
        finally
        {
            foreach (var path in paths) if (File.Exists(path)) File.Delete(path);
            Directory.Delete(directory);
        }
    }

    [Fact]
    public void IgnoresWordCommentRunReserializationButDetectsTextAndFontChanges()
    {
        const string ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        var directory = Path.Combine(Path.GetTempPath(), "nku-fingerprint-comments-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var paths = new[] { "before", "marked", "text-change", "font-change", "author-highlight-change" }
            .Select(name => Path.Combine(directory, name + ".docx")).ToArray();
        try
        {
            string before = $"<w:document {ns}><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/><w:highlight w:val=\"yellow\"/></w:rPr><w:t>前段 </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/></w:rPr><w:t>待修改的英文</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/></w:rPr><w:t> 后段</w:t></w:r></w:p></w:body></w:document>";
            string marked = $"<w:document {ns}><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/><w:highlight w:val=\"yellow\"/></w:rPr><w:t>前段 </w:t></w:r><w:bookmarkStart w:id=\"1\" w:name=\"_NKU_MARK_demo\"/><w:commentRangeStart w:id=\"2\"/><w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/><w:highlight w:val=\"yellow\"/></w:rPr><w:t>待修改</w:t><w:t>的英文</w:t></w:r><w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:commentReference w:id=\"2\"/></w:r><w:commentRangeEnd w:id=\"2\"/><w:bookmarkEnd w:id=\"1\"/><w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/></w:rPr><w:t> 后段</w:t></w:r></w:p></w:body></w:document>";
            WriteWordParts(paths[0], before, $"<w:settings {ns}><w:zoom w:percent=\"100\"/></w:settings>");
            WriteWordParts(paths[1], marked, $"<w:settings {ns}><w:zoom w:percent=\"100\"/></w:settings>");
            WriteWordParts(paths[2], marked.Replace("待修改</w:t>", "待改动</w:t>", StringComparison.Ordinal),
                $"<w:settings {ns}><w:zoom w:percent=\"100\"/></w:settings>");
            WriteWordParts(paths[3], marked.Replace("w:ascii=\"Arial\"", "w:ascii=\"Calibri\"", StringComparison.Ordinal),
                $"<w:settings {ns}><w:zoom w:percent=\"100\"/></w:settings>");
            WriteWordParts(paths[4], marked.Replace("w:highlight w:val=\"yellow\"/></w:rPr><w:t>前段", "w:highlight w:val=\"green\"/></w:rPr><w:t>前段", StringComparison.Ordinal),
                $"<w:settings {ns}><w:zoom w:percent=\"100\"/></w:settings>");

            Assert.True(DocxSnapshotFingerprint.Compute(paths[0]) == DocxSnapshotFingerprint.Compute(paths[1]),
                DocxSnapshotFingerprint.DescribeDifferences(paths[0], paths[1]));
            Assert.NotEqual(DocxSnapshotFingerprint.Compute(paths[0]), DocxSnapshotFingerprint.Compute(paths[2]));
            Assert.NotEqual(DocxSnapshotFingerprint.Compute(paths[0]), DocxSnapshotFingerprint.Compute(paths[3]));
            Assert.NotEqual(DocxSnapshotFingerprint.Compute(paths[0]), DocxSnapshotFingerprint.Compute(paths[4]));
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

    private static void WriteWordParts(string path, string documentXml, string settingsXml)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var (name, xml) in new[] { ("word/document.xml", documentXml), ("word/settings.xml", settingsXml) })
        {
            using var writer = new StreamWriter(archive.CreateEntry(name).Open());
            writer.Write(xml);
        }
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

    [Fact]
    public void IgnoresDerivedWordStatisticsButRetainsOtherApplicationProperties()
    {
        var directory=Path.Combine(Path.GetTempPath(),"nku-statistics-"+Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var paths=new[] { "first", "recounted", "edited" }.Select(x=>Path.Combine(directory,x+".docx")).ToArray();
        try {
            for(int index=0;index<paths.Length;index++) {
                WriteWordExportPackage(paths[index],"00112233","11223344");
                using var archive=ZipFile.Open(paths[index],ZipArchiveMode.Update);
                using var writer=new StreamWriter(archive.CreateEntry("docProps/app.xml").Open());
                writer.Write("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Company>"+(index==2 ? "edited" : "same")+"</Company><Pages>"+index+"</Pages><Words>"+(100+index)+"</Words><TotalTime>"+index+"</TotalTime></Properties>");
            }
            Assert.Equal(DocxSnapshotFingerprint.Compute(paths[0]),DocxSnapshotFingerprint.Compute(paths[1]));
            Assert.Empty(DocxSnapshotFingerprint.GetChangedParts(paths[0],paths[1]));
            Assert.NotEqual(DocxSnapshotFingerprint.Compute(paths[0]),DocxSnapshotFingerprint.Compute(paths[2]));
        } finally { foreach(var path in paths) File.Delete(path); Directory.Delete(directory); }
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

