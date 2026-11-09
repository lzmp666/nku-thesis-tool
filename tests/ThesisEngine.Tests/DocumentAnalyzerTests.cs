using System.IO.Compression;
using ThesisEngine.Core.Document;
using Xunit;

namespace ThesisEngine.Tests;

/// <summary>
/// #112 文档级版式检查。
///
/// 样本在 tests/fixtures/synthetic/document/，每个只偏离一个维度。
/// 除了样本，还手工造了几个样本覆盖不到的情形：NUMPAGES 不是页码、节的顺序、
/// 缺 body 级 sectPr、fldSimple 形式的域。
/// </summary>
public sealed class DocumentAnalyzerTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Fact]
    public void BaselineSampleHasNoIssues()
    {
        var result = Analyze("document_baseline_correct.docx");

        Assert.Single(result.Sections);
        Assert.Empty(result.Issues);
        Assert.False(result.EvenAndOddHeaders);
    }

    [Fact]
    public void LetterPaperIsReported()
    {
        var result = Analyze("document_paper_wrong.docx");

        var issue = Assert.Single(result.Issues);
        Assert.Equal("D1", issue.Code);
        Assert.Equal(1, issue.SectionIndex);
    }

    [Fact]
    public void WrongMarginsReportOneIssuePerSide()
    {
        var result = Analyze("document_margin_wrong.docx");

        Assert.Equal(2, result.Issues.Count(issue => issue.Code == "D2"));
        Assert.Contains(result.Issues, issue => issue.Title.Contains("上页边距"));
        Assert.Contains(result.Issues, issue => issue.Title.Contains("下页边距"));
        // 左右没改，不该跟着报。
        Assert.DoesNotContain(result.Issues, issue => issue.Title.Contains("左页边距"));
    }

    [Fact]
    public void WrongHeaderFooterDistanceReportsBothSides()
    {
        var result = Analyze("document_hf_distance_wrong.docx");

        Assert.Equal(2, result.Issues.Count(issue => issue.Code == "D3"));
    }

    [Fact]
    public void GutterIsReported()
    {
        var result = Analyze("document_gutter_wrong.docx");

        var issue = Assert.Single(result.Issues);
        Assert.Equal("D4", issue.Code);
    }

    /// <summary>整个文档没有页眉：报一条全文级，不要每节报一遍。</summary>
    [Fact]
    public void MissingHeaderIsReportedOnce()
    {
        var result = Analyze("document_no_header.docx");

        var issue = Assert.Single(result.Issues);
        Assert.Equal("D6", issue.Code);
        Assert.Equal(0, issue.SectionIndex);
    }

    /// <summary>
    /// 页眉引用了部件但里面没字：报一条全文级，不逐节报。
    /// 这条按全文判是有实测依据的——Word 保存时会为每节建好 default/even/first
    /// 三组空页眉页脚引用，逐节判会把封面节也报进去。
    /// </summary>
    [Fact]
    public void EmptyHeaderPartIsReportedOnceForWholeDocument()
    {
        var result = Analyze("document_empty_header.docx");

        var issue = Assert.Single(result.Issues);
        Assert.Equal("D6", issue.Code);
        Assert.Equal(0, issue.SectionIndex);
        Assert.Contains("第 1 节", issue.Current);
    }

    /// <summary>页脚是手打的 1，不是 PAGE 域。手打数字不会随后续页面变化。</summary>
    [Fact]
    public void PlainNumberInFooterIsNotAPageField()
    {
        var result = Analyze("document_footer_text.docx");

        var issue = Assert.Single(result.Issues);
        Assert.Equal("D7", issue.Code);
    }

    [Fact]
    public void EvenAndOddHeadersIsReported()
    {
        var result = Analyze("document_even_odd.docx");

        var issue = Assert.Single(result.Issues);
        Assert.Equal("D5", issue.Code);
        Assert.True(result.EvenAndOddHeaders);
    }

    [Fact]
    public void ManualTableOfContentsIsReported()
    {
        var result = Analyze("document_toc_manual.docx");

        var issue = Assert.Single(result.Issues);
        Assert.Equal("D9", issue.Code);
        Assert.False(result.HasTocField);
    }

    [Fact]
    public void AutomaticTableOfContentsIsAccepted()
    {
        var result = Analyze("document_toc_auto.docx");

        Assert.Empty(result.Issues);
        Assert.True(result.HasTocField);
    }

    /// <summary>空白页属惯例项，不能算成不符合规范。</summary>
    [Fact]
    public void BlankPageIsConventionOnly()
    {
        var result = Analyze("document_blank_page.docx");

        var issue = Assert.Single(result.Issues);
        Assert.Equal("D10", issue.Code);
        Assert.False(issue.HasClause);
        Assert.Empty(result.ClauseIssues);
    }

    [Fact]
    public void IdenticalNumberingAcrossSectionsIsReported()
    {
        var result = Analyze("document_sections_same_numbering.docx");

        Assert.Equal(2, result.Sections.Count);
        var issue = Assert.Single(result.Issues);
        Assert.Equal("D8", issue.Code);
    }

    [Fact]
    public void MixedNumberingAcrossSectionsIsAccepted()
    {
        var result = Analyze("document_sections_mixed_numbering.docx");

        Assert.Equal(2, result.Sections.Count);
        Assert.Empty(result.Issues);
        Assert.Equal("upperRoman", result.Sections[0].PageNumberFormat);
        Assert.Equal("decimal", result.Sections[1].PageNumberFormat);
    }

    /// <summary>
    /// 关键回归：NUMPAGES 是总页数，不是页码。只看「有没有 PAGE 字样」会把它当成页码域，
    /// 于是一份页脚只放了「第 X 页 共 Y 页」的文档被判为有页码，实际没有。
    /// </summary>
    [Fact]
    public void NumpagesFieldIsNotMistakenForPageNumber()
    {
        var path = CreateDocx(
            BodyWithSection(headerReference: true, footerReference: true),
            footerXml: $"""
                <w:ftr xmlns:w="{W}"><w:p>
                  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                  <w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
                  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                  <w:r><w:t>10</w:t></w:r>
                  <w:r><w:fldChar w:fldCharType="end"/></w:r>
                </w:p></w:ftr>
                """);
        try
        {
            var result = new DocumentAnalyzer().Analyze(path);

            Assert.Contains(result.Issues, issue => issue.Code == "D7");
        }
        finally { File.Delete(path); }
    }

    /// <summary>fldSimple 写法的 TOC 也要认出来，不能只认 instrText 三段式。</summary>
    [Fact]
    public void SimpleFieldTocIsRecognized()
    {
        var body = $"""
            <w:p><w:r><w:t>目录</w:t></w:r></w:p>
            <w:p><w:fldSimple w:instr=" TOC \o &quot;1-3&quot; \h "><w:r><w:t>第一章 引言</w:t></w:r></w:fldSimple></w:p>
            {SectionXml(true, true)}
            """;
        var path = CreateDocx(body);
        try
        {
            var result = new DocumentAnalyzer().Analyze(path);

            Assert.True(result.HasTocField);
            Assert.DoesNotContain(result.Issues, issue => issue.Code == "D9");
        }
        finally { File.Delete(path); }
    }

    /// <summary>节的顺序必须按文档顺序，分节符在段落里、最后一节在 body 末尾。</summary>
    [Fact]
    public void SectionOrderFollowsDocumentOrder()
    {
        var body = $"""
            <w:p><w:r><w:t>前置部分</w:t></w:r></w:p>
            <w:p><w:pPr>{SectionXml(true, true, "upperRoman")}</w:pPr></w:p>
            <w:p><w:r><w:t>正文</w:t></w:r></w:p>
            {SectionXml(true, true, "decimal")}
            """;
        var path = CreateDocx(body);
        try
        {
            var result = new DocumentAnalyzer().Analyze(path);

            Assert.Equal(2, result.Sections.Count);
            Assert.Equal("upperRoman", result.Sections[0].PageNumberFormat);
            Assert.Equal("decimal", result.Sections[1].PageNumberFormat);
        }
        finally { File.Delete(path); }
    }

    /// <summary>没有 body 级 sectPr 时，最后一节要补上，否则少报一节。</summary>
    [Fact]
    public void MissingBodySectionPropertiesFallsBackToLastSection()
    {
        var body = $"""
            <w:p><w:r><w:t>一节内容</w:t></w:r></w:p>
            <w:p><w:pPr>{SectionXml(true, true)}</w:pPr></w:p>
            <w:p><w:r><w:t>最后一节内容</w:t></w:r></w:p>
            """;
        var path = CreateDocx(body);
        try
        {
            var result = new DocumentAnalyzer().Analyze(path);

            Assert.Equal(2, result.Sections.Count);
        }
        finally { File.Delete(path); }
    }

    /// <summary>
    /// 关键回归：多节文档里封面节页眉为空、正文节有页眉，不该报。
    ///
    /// 这是 Word 实测逼出来的口径。Word 保存时会为每节建好 default/even/first 三组
    /// 页眉页脚引用，即使内容是空的；而规范说「篇眉从中文摘要开始」，封面本来就没页眉。
    /// 若按节判「页眉为空」，每篇合格论文都会被报一遍。
    /// </summary>
    [Fact]
    public void CoverSectionWithoutHeaderIsNotReportedWhenBodySectionHasOne()
    {
        var body = $"""
            <w:p><w:r><w:t>封面</w:t></w:r></w:p>
            <w:p><w:pPr><w:sectPr>
              <w:headerReference w:type="default" r:id="rId12"/>
              <w:pgSz w:w="11906" w:h="16838"/>
              <w:pgMar w:top="2155" w:right="1814" w:bottom="2155" w:left="1814" w:header="1701" w:footer="1701" w:gutter="0"/>
            </w:sectPr></w:pPr></w:p>
            <w:p><w:r><w:t>正文</w:t></w:r></w:p>
            {SectionXml(true, true)}
            """;
        var path = CreateDocx(body, extraEmptyHeader: true);
        try
        {
            var result = new DocumentAnalyzer().Analyze(path);

            Assert.Equal(2, result.Sections.Count);
            Assert.DoesNotContain(result.Issues, issue => issue.Code == "D6");
        }
        finally { File.Delete(path); }
    }

    private static DocumentAnalysisResult Analyze(string sampleName)
    {
        var root = FindRepositoryRoot();
        var path = Path.Combine(root, "tests", "fixtures", "synthetic", "document", sampleName);
        Assert.True(File.Exists(path), $"样本不存在：{path}");
        return new DocumentAnalyzer().Analyze(path);
    }

    private static string BodyWithSection(bool headerReference, bool footerReference) =>
        $"<w:p><w:r><w:t>正文</w:t></w:r></w:p>{SectionXml(headerReference, footerReference)}";

    private static string SectionXml(bool headerReference, bool footerReference, string? numberFormat = null)
    {
        var parts = new List<string>();
        if (headerReference) parts.Add($"<w:headerReference w:type=\"default\" r:id=\"rId10\"/>");
        if (footerReference) parts.Add($"<w:footerReference w:type=\"default\" r:id=\"rId11\"/>");
        parts.Add("<w:pgSz w:w=\"11906\" w:h=\"16838\"/>");
        parts.Add("<w:pgMar w:top=\"2155\" w:right=\"1814\" w:bottom=\"2155\" w:left=\"1814\" " +
            "w:header=\"1701\" w:footer=\"1701\" w:gutter=\"0\"/>");
        if (numberFormat != null) parts.Add($"<w:pgNumType w:fmt=\"{numberFormat}\" w:start=\"1\"/>");
        return "<w:sectPr>" + string.Concat(parts) + "</w:sectPr>";
    }

    private static string CreateDocx(string bodyXml, string? footerXml = null, bool extraEmptyHeader = false)
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

        Add(archive, "[Content_Types].xml",
            $$"""
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="xml" ContentType="application/xml"/>
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
              <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
              <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
              <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
            </Types>
            """);
        Add(archive, "_rels/.rels",
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""");
        Add(archive, "word/document.xml",
            $"""<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{bodyXml}</w:body></w:document>""");
        Add(archive, "word/_rels/document.xml.rels",
            $"""
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
              <Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
              <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
            </Relationships>
            """);
        Add(archive, "word/header1.xml",
            $"""<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>第一章 引言</w:t></w:r></w:p></w:hdr>""");
        if (extraEmptyHeader)
        {
            Add(archive, "word/header2.xml",
                $"""<w:hdr xmlns:w="{W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p></w:hdr>""");
        }
        Add(archive, "word/footer1.xml", footerXml ?? PageFieldFooter());

        return path;
    }

    private static string PageFieldFooter() =>
        $"""
        <w:ftr xmlns:w="{W}"><w:p>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>1</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p></w:ftr>
        """;

    private static void Add(ZipArchive archive, string name, string content)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open());
        writer.Write(content);
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }
}
