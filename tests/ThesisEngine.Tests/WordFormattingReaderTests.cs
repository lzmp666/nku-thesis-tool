using System.IO.Compression;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class WordFormattingReaderTests
{
    [Fact]
    public void ReadsDirectFontsAndHalfPointSizeWithoutChangingThePackage()
    {
        var path = CreatePackage(Document(
            "<w:p><w:r><w:rPr><w:rFonts w:eastAsia=\"中文字体\" w:ascii=\"Latin Font\"/><w:sz w:val=\"21\"/></w:rPr><w:t>文本</w:t></w:r></w:p>"));
        try
        {
            var before = File.ReadAllBytes(path);
            var snapshot = new WordFormattingReader().Read(path);
            var run = Assert.Single(Assert.Single(snapshot.Paragraphs).Runs);

            Assert.Equal("中文字体", run.EastAsiaFont);
            Assert.Equal("Latin Font", run.AsciiFont);
            Assert.Equal(10.5, run.FontSizePt);
            Assert.Equal("文本", run.Text);
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void ResolvesDocDefaultsBasedOnStylesAndDirectRunOverride()
    {
        var styles = """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/><w:rFonts w:eastAsia="Default East"/></w:rPr></w:rPrDefault></w:docDefaults>
              <w:style w:type="paragraph" w:styleId="Parent"><w:rPr><w:rFonts w:ascii="Parent Latin"/></w:rPr></w:style>
              <w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Parent"/><w:rPr><w:rFonts w:eastAsia="Style East"/></w:rPr></w:style>
            </w:styles>
            """;
        var path = CreatePackage(Document(
            "<w:p><w:pPr><w:pStyle w:val=\"Child\"/></w:pPr><w:r><w:rPr><w:sz w:val=\"22\"/></w:rPr><w:t>覆盖</w:t></w:r></w:p>"), styles);
        try
        {
            var run = Assert.Single(Assert.Single(new WordFormattingReader().Read(path).Paragraphs).Runs);
            Assert.Equal("Style East", run.EastAsiaFont);
            Assert.Equal("Parent Latin", run.AsciiFont);
            Assert.Equal(11, run.FontSizePt);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void ParagraphMarkFormattingDoesNotOverrideTextRunFormatting()
    {
        var styles = """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia="正文默认字体"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
            </w:styles>
            """;
        var path = CreatePackage(Document(
            "<w:p><w:pPr><w:rPr><w:rFonts w:eastAsia=\"段落标记字体\"/><w:sz w:val=\"48\"/></w:rPr></w:pPr><w:r><w:t>正文</w:t></w:r></w:p>"),
            styles);
        try
        {
            var run = Assert.Single(Assert.Single(new WordFormattingReader().Read(path).Paragraphs).Runs);
            Assert.Equal("正文默认字体", run.EastAsiaFont);
            Assert.Equal(12, run.FontSizePt);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void MarksWrappedAndNestedBodyContentAsUnsupported()
    {
        var path = CreatePackage(Document(
            "<w:p><w:r><w:t>正文</w:t></w:r><w:hyperlink><w:r><w:t>链接</w:t></w:r></w:hyperlink><w:ins><w:r><w:t>修订</w:t></w:r></w:ins></w:p>" +
            "<w:sdt><w:sdtContent><w:p><w:r><w:t>内容控件</w:t></w:r></w:p></w:sdtContent></w:sdt>"));
        try
        {
            var snapshot = new WordFormattingReader().Read(path);
            Assert.Single(snapshot.Paragraphs);
            Assert.Single(snapshot.Paragraphs[0].Runs);
            Assert.Contains("hyperlinks", snapshot.UnsupportedScopes);
            Assert.Contains("tracked changes", snapshot.UnsupportedScopes);
            Assert.Contains("content controls", snapshot.UnsupportedScopes);
            Assert.Contains("nested body paragraphs", snapshot.UnsupportedScopes);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void KeepsRunsIndependentAndMarksUnsupportedScopes()
    {
        var path = CreatePackage(Document(
            "<w:p><w:r><w:rPr><w:sz w:val=\"24\"/></w:rPr><w:t>一</w:t></w:r><w:r><w:t>二</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>表</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"),
            entries: new[]
            {
                ("word/header1.xml", "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>"),
                ("word/footnotes.xml", "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>"),
                ("word/endnotes.xml", "<w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>")
            });
        try
        {
            var snapshot = new WordFormattingReader().Read(path);
            var runs = Assert.Single(snapshot.Paragraphs).Runs;
            Assert.Equal(2, runs.Count);
            Assert.Equal(0, runs[0].RunIndex);
            Assert.Equal(1, runs[1].RunIndex);
            Assert.Equal(12, runs[0].FontSizePt);
            Assert.Null(runs[1].FontSizePt);
            Assert.Contains("tables", snapshot.UnsupportedScopes);
            Assert.Contains("headers/footers", snapshot.UnsupportedScopes);
            Assert.Contains("footnotes", snapshot.UnsupportedScopes);
            Assert.Contains("endnotes", snapshot.UnsupportedScopes);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void MissingThemeIsUnknownAndExplicitZeroDoesNotInherit()
    {
        var styles = """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
            </w:styles>
            """;
        var path = CreatePackage(Document(
            "<w:p><w:r><w:rPr><w:rFonts w:eastAsiaTheme=\"minorEastAsia\"/><w:sz w:val=\"0\"/></w:rPr><w:t>未知</w:t></w:r></w:p>"), styles);
        try
        {
            var snapshot = new WordFormattingReader().Read(path);
            var run = Assert.Single(Assert.Single(snapshot.Paragraphs).Runs);
            Assert.Null(run.EastAsiaFont);
            Assert.Null(run.FontSizePt);
            Assert.Contains("word/theme/theme1.xml", snapshot.UnsupportedScopes);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void ResolvesEastAsiaThemeFontFromDrawingMlTheme()
    {
        var theme = """
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:themeElements><a:fontScheme name="Test"><a:majorFont><a:latin typeface="Latin"/><a:ea typeface=""/><a:font script="Hans" typeface="主题中文"/></a:majorFont><a:minorFont><a:latin typeface="Minor Latin"/><a:ea typeface=""/></a:minorFont></a:fontScheme></a:themeElements>
            </a:theme>
            """;
        var styles = """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:lang w:eastAsia="zh-CN"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>
            """;
        var path = CreatePackage(Document("<w:p><w:r><w:rPr><w:rFonts w:eastAsiaTheme=\"majorEastAsia\"/></w:rPr><w:t>主题</w:t></w:r></w:p>"), styles,
            new[] { ("word/theme/theme1.xml", theme) });
        try
        {
            var run = Assert.Single(Assert.Single(new WordFormattingReader().Read(path).Paragraphs).Runs);
            Assert.Equal("主题中文", run.EastAsiaFont);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DoesNotInheritBasedOnAcrossStyleTypes()
    {
        var styles = """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:style w:type="character" w:styleId="CharacterParent"><w:rPr><w:sz w:val="24"/></w:rPr></w:style>
              <w:style w:type="paragraph" w:styleId="ParagraphChild"><w:basedOn w:val="CharacterParent"/></w:style>
            </w:styles>
            """;
        var path = CreatePackage(Document(
            "<w:p><w:pPr><w:pStyle w:val=\"ParagraphChild\"/></w:pPr><w:r><w:t>类型</w:t></w:r></w:p>"), styles);
        try
        {
            var run = Assert.Single(Assert.Single(new WordFormattingReader().Read(path).Paragraphs).Runs);
            Assert.Null(run.FontSizePt);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void AppliesCharacterStyleAfterParagraphStyle()
    {
        var styles = """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:style w:type="paragraph" w:styleId="Paragraph"><w:rPr><w:sz w:val="24"/></w:rPr></w:style>
              <w:style w:type="character" w:styleId="Character"><w:rPr><w:sz w:val="21"/></w:rPr></w:style>
            </w:styles>
            """;
        var path = CreatePackage(Document(
            "<w:p><w:pPr><w:pStyle w:val=\"Paragraph\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"Character\"/></w:rPr><w:t>字符样式</w:t></w:r></w:p>"), styles);
        try
        {
            var run = Assert.Single(Assert.Single(new WordFormattingReader().Read(path).Paragraphs).Runs);
            Assert.Equal(10.5, run.FontSizePt);
        }
        finally { File.Delete(path); }
    }

    private static string CreatePackage(string document, string? styles = null, IEnumerable<(string Name, string Xml)>? entries = null)
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "word/document.xml", document);
        if (styles is not null) Add(archive, "word/styles.xml", styles);
        foreach (var entry in entries ?? Array.Empty<(string Name, string Xml)>()) Add(archive, entry.Name, entry.Xml);
        return path;
    }

    private static string Document(string body, string? styles = null) =>
        $"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{body}<w:sectPr/></w:body></w:document>";

    private static void Add(ZipArchive archive, string name, string xml)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open());
        writer.Write(xml);
    }
}
