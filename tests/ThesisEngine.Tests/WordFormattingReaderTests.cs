using System.IO.Compression;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class WordFormattingReaderTests
{
    [Fact]
    public void InheritsRenamedDefaultParagraphStyleWhenWordOmitsPStyle()
    {
        var styles = """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
              <w:style w:type="paragraph" w:default="1" w:styleId="a1">
                <w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr>
                <w:rPr><w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>
              </w:style>
            </w:styles>
            """;
        var path = CreatePackage(Document(
            "<w:p><w:r><w:t>正常正文 English</w:t></w:r></w:p>" +
            "<w:p><w:r><w:rPr><w:rFonts w:eastAsia=\"黑体\"/><w:sz w:val=\"21\"/></w:rPr><w:t>直接格式覆盖</w:t></w:r></w:p>"), styles);
        try
        {
            var snapshot = new WordFormattingReader().Read(path);
            var inherited = Assert.Single(snapshot.Paragraphs[0].Runs);
            Assert.Equal("宋体", inherited.EastAsiaFont);
            Assert.Equal("Times New Roman", inherited.AsciiSlotFont);
            Assert.Equal(12, inherited.FontSizePt);
            var direct = Assert.Single(snapshot.Paragraphs[1].Runs);
            Assert.Equal("黑体", direct.EastAsiaFont);
            Assert.Equal(10.5, direct.FontSizePt);
        }
        finally { File.Delete(path); }
    }

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
            Assert.Equal(2, snapshot.Paragraphs.Count);
            Assert.Single(snapshot.Paragraphs[0].Runs);
            Assert.Single(snapshot.Paragraphs[1].Runs);
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
            Assert.Equal(2, snapshot.Paragraphs.Count);
            var runs = snapshot.Paragraphs[0].Runs;
            Assert.Equal(2, runs.Count);
            Assert.Equal(0, runs[0].RunIndex);
            Assert.Equal(1, runs[1].RunIndex);
            Assert.Equal(12, runs[0].FontSizePt);
            Assert.Null(runs[1].FontSizePt);
            Assert.Equal(ParagraphScopeClassification.Excluded, snapshot.Paragraphs[1].Scope);
            Assert.Contains("表格单元格", snapshot.Paragraphs[1].ScopeReason);
            Assert.Contains(snapshot.UnsupportedScopes, scope => scope.StartsWith("tables", StringComparison.Ordinal));
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

    [Fact]
    public void ConfirmsOnlyParagraphsBetweenRecognizedChapterAndNonBodyHeadings()
    {
        var styles = """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
              <w:style w:type="paragraph" w:styleId="Heading2"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
            </w:styles>
            """;
        var body =
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>摘要</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>摘要文字</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>第一章绪论</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>正文段落</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:pStyle w:val=\"Heading2\"/></w:pPr><w:r><w:t>研究方法</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>第二个正文段落</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>参考文献</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>文献条目</w:t></w:r></w:p>";
        var path = CreatePackage(Document(body), styles);
        try
        {
            var paragraphs = new WordFormattingReader().Read(path).Paragraphs;

            Assert.Equal(8, paragraphs.Count);
            Assert.Equal(ParagraphScopeClassification.Excluded, paragraphs[0].Scope);
            Assert.Equal(ParagraphScopeClassification.Unknown, paragraphs[1].Scope);
            Assert.Equal(ParagraphScopeClassification.Excluded, paragraphs[2].Scope);
            Assert.Equal(ParagraphScopeClassification.ConfirmedBodyProse, paragraphs[3].Scope);
            Assert.Equal(ParagraphScopeClassification.Excluded, paragraphs[4].Scope);
            Assert.Equal(ParagraphScopeClassification.ConfirmedBodyProse, paragraphs[5].Scope);
            Assert.Equal(ParagraphScopeClassification.Excluded, paragraphs[6].Scope);
            Assert.Equal(ParagraphScopeClassification.Unknown, paragraphs[7].Scope);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void LeavesParagraphsUnknownWhenEitherBoundaryCannotBeRecognized()
    {
        var path = CreatePackage(Document("<w:p><w:r><w:t>普通段落</w:t></w:r></w:p>"));
        try
        {
            var paragraph = Assert.Single(new WordFormattingReader().Read(path).Paragraphs);
            Assert.Equal(ParagraphScopeClassification.Unknown, paragraph.Scope);
            Assert.Contains("未同时识别", paragraph.ScopeReason);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void FlattensTableParagraphsForStableLocationsButExcludesThemFromBodyRules()
    {
        var body =
            "<w:p><w:r><w:t>正文段落</w:t></w:r></w:p>" +
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格单元格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "<w:p><w:pPr><w:pStyle w:val=\"Caption\"/></w:pPr><w:r><w:t>表1-1 样本统计</w:t></w:r></w:p>" +
            "<w:p><w:r><w:drawing/></w:r></w:p>";
        var path = CreatePackage(Document(body));
        try
        {
            var snapshot = new WordFormattingReader().Read(path);

            Assert.Equal(4, snapshot.Paragraphs.Count);
            Assert.Equal(0, snapshot.Paragraphs[0].ParagraphIndex);
            Assert.Equal(1, snapshot.Paragraphs[1].ParagraphIndex);
            Assert.Equal(2, snapshot.Paragraphs[2].ParagraphIndex);
            Assert.Equal(3, snapshot.Paragraphs[3].ParagraphIndex);
            Assert.Equal(ParagraphScopeClassification.Excluded, snapshot.Paragraphs[1].Scope);
            Assert.Contains("表格单元格", snapshot.Paragraphs[1].ScopeReason);
            Assert.Equal(ParagraphScopeClassification.Excluded, snapshot.Paragraphs[2].Scope);
            Assert.Contains("题注专属规则", snapshot.Paragraphs[2].ScopeReason);
            Assert.Contains(snapshot.UnsupportedScopes, scope => scope.StartsWith("tables", StringComparison.Ordinal));
            Assert.Contains(snapshot.UnsupportedScopes, scope => scope.StartsWith("images", StringComparison.Ordinal));
            Assert.Equal(CaptionKind.Table, snapshot.Paragraphs[2].Caption);
            Assert.False(snapshot.Paragraphs[2].CaptionPlacementCorrect);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void RecognizesStyledEnglishChapterAndChecksItsLatinText()
    {
        var styles = """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
            </w:styles>
            """;
        var body =
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>Chapter 1 Introduction</w:t></w:r></w:p>" +
            "<w:p><w:r><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr><w:t>This is a synthetic English thesis paragraph.</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>References</w:t></w:r></w:p>";
        var path = CreatePackage(Document(body), styles);
        try
        {
            var snapshot = new WordFormattingReader().Read(path);
            var bodyParagraph = snapshot.Paragraphs[1];
            var analysis = new NkuGraduate2026BodyAnalyzer().Analyze(snapshot, NkuGraduate2026BodyProfile.Draft);

            Assert.Equal(ParagraphScopeClassification.ConfirmedBodyProse, bodyParagraph.Scope);
            Assert.Collection(analysis.Diagnostics,
                diagnostic => Assert.Equal(NkuGraduate2026BodyAnalyzer.RuleIdLatinFont, diagnostic.RuleId));
            Assert.Empty(analysis.UnsupportedScopes);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DoesNotTreatA正文SentenceMentioningATableAsTheCaption()
    {
        var body =
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>第一章 绪论</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>表4.1说明了合成样本的区域指标差异。</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>参考文献</w:t></w:r></w:p>";
        var path = CreatePackage(Document(body));
        try
        {
            var paragraph = new WordFormattingReader().Read(path).Paragraphs[1];

            Assert.Null(paragraph.Caption);
            Assert.Equal(ParagraphScopeClassification.ConfirmedBodyProse, paragraph.Scope);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void KeepsAsciiAndHAnsiSlotsIndependentForThe2026Profile()
    {
        var path = CreatePackage(Document(
            "<w:p><w:r><w:rPr><w:rFonts w:eastAsia=\"宋体\" w:ascii=\"Times New Roman\" w:hAnsi=\"Arial\"/><w:sz w:val=\"24\"/></w:rPr><w:t>中文 Latin</w:t></w:r></w:p>"));
        try
        {
            var run = Assert.Single(Assert.Single(new WordFormattingReader().Read(path).Paragraphs).Runs);

            Assert.Equal("宋体", run.EastAsiaFont);
            Assert.Equal("Times New Roman", run.AsciiFont);
            Assert.Equal("Times New Roman", run.AsciiSlotFont);
            Assert.Equal("Arial", run.HAnsiFont);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void InheritsAsciiAndHAnsiFontSlotsIndependentlyAcrossStyles()
    {
        var styles = """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Default ASCII" w:hAnsi="Default HAnsi"/></w:rPr></w:rPrDefault></w:docDefaults>
              <w:style w:type="paragraph" w:styleId="Parent"><w:rPr><w:rFonts w:ascii="Parent ASCII"/></w:rPr></w:style>
              <w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Parent"/><w:rPr><w:rFonts w:hAnsi="Child HAnsi"/></w:rPr></w:style>
            </w:styles>
            """;
        var path = CreatePackage(Document(
            "<w:p><w:pPr><w:pStyle w:val=\"Child\"/></w:pPr><w:r><w:t>Latin</w:t></w:r></w:p>"), styles);
        try
        {
            var run = Assert.Single(Assert.Single(new WordFormattingReader().Read(path).Paragraphs).Runs);

            Assert.Equal("Parent ASCII", run.AsciiSlotFont);
            Assert.Equal("Child HAnsi", run.HAnsiFont);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void LeavesParagraphWithUnsupportedInlineContentOutsideConfirmedBodyScope()
    {
        var styles = """
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
            </w:styles>
            """;
        var body =
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>第一章 绪论</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>普通正文</w:t></w:r><w:hyperlink><w:r><w:t>链接文本</w:t></w:r></w:hyperlink></w:p>" +
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>参考文献</w:t></w:r></w:p>";
        var path = CreatePackage(Document(body), styles);
        try
        {
            var snapshot = new WordFormattingReader().Read(path);

            Assert.Equal(ParagraphScopeClassification.Unknown, snapshot.Paragraphs[1].Scope);
            Assert.Contains("hyperlinks", snapshot.UnsupportedScopes);
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
