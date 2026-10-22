using System.IO.Compression;
using System.Xml.Linq;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class WordAnnotationWriterTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";

    [Fact]
    public void GroupedDiagnosticCreatesOneCommentAndHighlightsEveryOccurrence()
    {
        var root = FindRepositoryRoot();
        var inputPath = Path.Combine(root.FullName, "tests", "fixtures", "synthetic", "demo", "known_issues_sample.docx");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            var diagnostic = new Diagnostic("demo/east-asian-font", "warning",
                "word/document.xml#paragraph[1]/run[0]", "宋体", "黑体、雅黑", "auto", "统一改为宋体。")
            {
                Occurrences = new[]
                {
                    new DiagnosticOccurrence("word/document.xml#paragraph[1]/run[0]", "黑体"),
                    new DiagnosticOccurrence("word/document.xml#paragraph[2]/run[0]", "雅黑")
                }
            };

            WordAnnotationWriter.WriteAnnotatedCopy(inputPath, outputPath, new[] { diagnostic });

            using var archive = ZipFile.OpenRead(outputPath);
            var document = ReadXml(archive, "word/document.xml");
            var paragraphs = document.Descendants(W + "body").Single().Elements(W + "p").ToArray();
            Assert.Equal("yellow", (string?)paragraphs[1].Descendants(W + "highlight").Single().Attribute(W + "val"));
            Assert.Equal("yellow", (string?)paragraphs[2].Descendants(W + "highlight").Single().Attribute(W + "val"));
            Assert.Single(document.Descendants(W + "commentRangeStart"));
            Assert.Single(document.Descendants(W + "commentRangeEnd"));

            var comments = ReadXml(archive, "word/comments.xml");
            var comment = Assert.Single(comments.Root!.Elements(W + "comment"));
            var hidden = string.Concat(comment.Descendants(W + "r")
                .Where(run => run.Element(W + "rPr")?.Element(W + "vanish") is not null)
                .SelectMany(run => run.Elements(W + "t").Select(text => text.Value)));
            Assert.Contains("[NKU-OCCURRENCE:word/document.xml#paragraph[1]/run[0]|absent]", hidden);
            Assert.Contains("[NKU-OCCURRENCE:word/document.xml#paragraph[2]/run[0]|absent]", hidden);
        }
        finally
        {
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void CommentShowsApprovedSourceAndHumanReadableRuleName()
    {
        var root = FindRepositoryRoot();
        var inputPath = Path.Combine(root.FullName, "tests", "fixtures", "synthetic", "demo", "known_issues_sample.docx");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        const string source = "南开大学研究生学位论文写作规范（2026版）第4.7节表4.6（M26-036）";
        try
        {
            var diagnostic = new Diagnostic("NKU-REQ-042", "warning",
                "word/document.xml#paragraph[1]/run[0]", "宋体", "黑体", "manual",
                source + "：中文字体应为「宋体」，实际为「黑体」。");

            WordAnnotationWriter.WriteAnnotatedCopy(inputPath, outputPath, new[] { diagnostic });

            using var archive = ZipFile.OpenRead(outputPath);
            var comments = ReadXml(archive, "word/comments.xml");
            var comment = Assert.Single(comments.Root!.Elements(W + "comment"));
            var visibleText = string.Concat(comment.Descendants(W + "r")
                .Where(run => run.Element(W + "rPr")?.Element(W + "vanish") is null)
                .SelectMany(run => run.Elements(W + "t").Select(text => text.Value)));
            Assert.Equal("要求参考来源：" + source + "；规则：中文字体；期望值：宋体；实际值：黑体。", visibleText);
            Assert.DoesNotContain("NKU-REQ-042", visibleText);
        }
        finally
        {
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void AnnotatesExistingSyntheticDemoFixture()
    {
        var root = FindRepositoryRoot();
        var inputPath = Path.Combine(root.FullName, "tests", "fixtures", "synthetic", "demo", "known_issues_sample.docx");
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            var originalBytes = File.ReadAllBytes(inputPath);
            var originalText = ReadTextRuns(inputPath);
            var diagnostics = new[]
            {
                new Diagnostic("demo/east-asian-font", "warning", "word/document.xml#paragraph[1]/run[0]", "宋体", "黑体", "auto", "请改为宋体。"),
                new Diagnostic("demo/western-font", "warning", "word/document.xml#paragraph[2]/run[0]", "Times New Roman", "Arial", "auto", "请改为 Times New Roman。"),
                new Diagnostic("demo/font-size", "warning", "word/document.xml#paragraph[3]/run[0]", "12pt", "10.5pt", "auto", "请调整字号。")
            };

            WordAnnotationWriter.WriteAnnotatedCopy(inputPath, outputPath, diagnostics);

            Assert.Equal(originalBytes, File.ReadAllBytes(inputPath));
            Assert.Equal(originalText, ReadTextRuns(outputPath));
            var rechecked = new FontSizeAnalyzer().Analyze(
                new WordFormattingReader().Read(outputPath),
                DemoRuleSet.NkuDraftSample);
            Assert.Equal(3, rechecked.Diagnostics.Count);
            Assert.DoesNotContain(rechecked.Diagnostics, diagnostic => diagnostic.Actual == "(unknown)");
            using var archive = ZipFile.OpenRead(outputPath);
            var document = ReadXml(archive, "word/document.xml");
            var body = document.Descendants(W + "body").Single();
            var paragraphs = body.Elements(W + "p").ToArray();
            Assert.Equal(5, paragraphs.Length);
            Assert.All(paragraphs.Skip(1).Take(3), paragraph =>
            {
                Assert.Single(paragraph.Elements(W + "commentRangeStart"));
                Assert.Single(paragraph.Elements(W + "commentRangeEnd"));
                Assert.Single(paragraph.Descendants(W + "commentReference"));
                var targetRun = paragraph.Elements(W + "r").Single(run => run.Descendants(W + "t").Any());
                var paragraphElements = paragraph.Elements().ToArray();
                var runPosition = Array.IndexOf(paragraphElements, targetRun);
                Assert.True(runPosition > 0 && runPosition + 2 < paragraphElements.Length);
                var rangeStart = paragraphElements[runPosition - 1];
                var rangeEnd = paragraphElements[runPosition + 1];
                var commentReference = paragraphElements[runPosition + 2];
                Assert.Equal(W + "commentRangeStart", rangeStart.Name);
                Assert.Equal(W + "commentRangeEnd", rangeEnd.Name);
                Assert.Equal((string?)rangeStart.Attribute(W + "id"), (string?)rangeEnd.Attribute(W + "id"));
                Assert.Equal((string?)rangeStart.Attribute(W + "id"),
                    (string?)commentReference.Element(W + "commentReference")?.Attribute(W + "id"));
                Assert.Equal("yellow", (string?)targetRun.Element(W + "rPr")?.Element(W + "highlight")?.Attribute(W + "val"));
            });

            var comments = ReadXml(archive, "word/comments.xml");
            Assert.Equal(3, comments.Root!.Elements(W + "comment").Count());
            var commentElements = comments.Root!.Elements(W + "comment").ToArray();
            var commentTexts = commentElements
                .Select(comment => string.Concat(comment.Descendants(W + "r")
                    .Where(run => run.Element(W + "rPr")?.Element(W + "vanish") is null)
                    .SelectMany(run => run.Elements(W + "t").Select(element => element.Value))))
                .ToArray();
            Assert.Equal(3, commentTexts.Length);
            Assert.Contains("要求参考来源：演示规则（仅用于验证插件流程，不代表正式标准）；规则：中文字体；期望值：宋体；实际值：黑体。", commentTexts);
            Assert.Contains("要求参考来源：演示规则（仅用于验证插件流程，不代表正式标准）；规则：西文字体；期望值：Times New Roman；实际值：Arial。", commentTexts);
            Assert.Contains("要求参考来源：演示规则（仅用于验证插件流程，不代表正式标准）；规则：字号；期望值：12pt；实际值：10.5pt。", commentTexts);
            Assert.DoesNotContain(commentTexts, text => text.Contains("demo/", StringComparison.Ordinal));
            var hiddenLocations = commentElements
                .SelectMany(comment => comment.Descendants(W + "r")
                    .Where(run => run.Element(W + "rPr")?.Element(W + "vanish") is not null)
                    .SelectMany(run => run.Elements(W + "t").Select(element => element.Value)))
                .ToArray();
            Assert.Contains("[NKU-LOCATION:word/document.xml#paragraph[1]/run[0]]", hiddenLocations);
            Assert.Contains("[NKU-TOOL-ANNOTATION:v1]", hiddenLocations);
            Assert.Contains("[NKU-HIGHLIGHT:absent]", hiddenLocations);
            Assert.DoesNotContain("NKU-LOCATION", string.Join("", commentTexts));
            Assert.All(commentTexts, text =>
            {
                Assert.DoesNotContain("修改建议：", text);
                Assert.DoesNotContain("检测提示：", text);
            });

            var relationships = ReadXml(archive, "word/_rels/document.xml.rels");
            Assert.Contains(relationships.Root!.Elements(Rel + "Relationship"), element =>
                (string?)element.Attribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" &&
                (string?)element.Attribute("Target") == "comments.xml");
        }
        finally
        {
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void AnnotatesTargetRunAndPreservesOriginalPackageContent()
    {
        var inputPath = CreatePackage();
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            var originalBytes = File.ReadAllBytes(inputPath);
            var result = WordAnnotationWriter.WriteAnnotatedCopy(inputPath, outputPath,
            [
                new Diagnostic("demo/east-asian-font", "warning", "word/document.xml#paragraph[0]/run[0]", "宋体", "黑体", "auto", "请改为宋体。"),
                new Diagnostic("demo/font-size", "warning", "word/document.xml#paragraph[0]/run[0]", "12pt", "14pt", "auto", "请调整字号。")
            ]);

            Assert.Equal(Path.GetFullPath(outputPath), result);
            Assert.Equal(originalBytes, File.ReadAllBytes(inputPath));

            using var archive = ZipFile.OpenRead(outputPath);
            var document = ReadXml(archive, "word/document.xml");
            var run = Assert.Single(document.Descendants(W + "r"), element => element.Descendants(W + "t").Any());
            var properties = run.Element(W + "rPr")!;
            Assert.Equal("yellow", (string?)properties.Element(W + "highlight")?.Attribute(W + "val"));
            Assert.Equal("Arial", (string?)properties.Element(W + "rFonts")?.Attribute(W + "ascii"));
            Assert.Equal("28", (string?)properties.Element(W + "sz")?.Attribute(W + "val"));
            Assert.Equal("single", (string?)properties.Element(W + "u")?.Attribute(W + "val"));
            Assert.Equal("原文内容", string.Concat(run.Descendants(W + "t").Select(text => text.Value)));

            var paragraph = document.Descendants(W + "body").Single().Elements(W + "p").Single();
            Assert.Equal(new[] { "1", "2" }, paragraph.Elements(W + "commentRangeStart")
                .Select(element => (string?)element.Attribute(W + "id")).ToArray());
            Assert.Equal(new[] { "2", "1" }, paragraph.Elements(W + "commentRangeEnd")
                .Select(element => (string?)element.Attribute(W + "id")).ToArray());
            Assert.Equal(new[] { "1", "2" }, paragraph.Descendants(W + "commentReference")
                .Select(element => (string?)element.Attribute(W + "id")).ToArray());

            var comments = ReadXml(archive, "word/comments.xml");
            Assert.Equal(new[] { "0", "1", "2" }, comments.Root!.Elements(W + "comment")
                .Select(element => (string?)element.Attribute(W + "id")).ToArray());
            var commentText = string.Join(" ", comments.Descendants(W + "t").Select(element => element.Value));
            Assert.Contains("要求参考来源：演示规则（仅用于验证插件流程，不代表正式标准）", commentText);
            Assert.Contains("规则：中文字体", commentText);
            Assert.DoesNotContain("2026版第4.7节表4.6", commentText);
            Assert.Contains("期望值：宋体；实际值：黑体", commentText);
            Assert.DoesNotContain("修改建议：", commentText);
            Assert.DoesNotContain("检测提示：", commentText);
            Assert.Contains("期望值：12pt；实际值：14pt", commentText);
            Assert.Contains("Existing note", commentText);
            Assert.Contains("[NKU-TOOL-ANNOTATION:v1]", commentText);
            Assert.Contains("[NKU-HIGHLIGHT:pink]", commentText);

            var relationships = ReadXml(archive, "word/_rels/document.xml.rels");
            Assert.Single(relationships.Root!.Elements(Rel + "Relationship"), element =>
                (string?)element.Attribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments");
            var contentTypes = ReadXml(archive, "[Content_Types].xml");
            Assert.Contains(contentTypes.Root!.Elements(Ct + "Override"), element =>
                (string?)element.Attribute("PartName") == "/word/comments.xml" &&
                (string?)element.Attribute("ContentType") == "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml");
            Assert.Equal("preserve-me", ReadEntry(archive, "custom/data.bin"));
        }
        finally
        {
            File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void RefreshRemovesOnlyToolOwnedCommentsAndRestoresOriginalHighlight()
    {
        var inputPath = CreatePackage();
        var firstOutput = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        var refreshedOutput = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        var diagnostic = new Diagnostic("demo/font-size", "warning", "word/document.xml#paragraph[0]/run[0]", "12pt", "14pt", "auto", "请调整字号。");
        try
        {
            WordAnnotationWriter.WriteAnnotatedCopy(inputPath, firstOutput, [diagnostic, diagnostic with { RuleId = "demo/east-asian-font" }]);
            WordAnnotationWriter.WriteAnnotatedCopy(firstOutput, refreshedOutput, Array.Empty<Diagnostic>());

            using var archive = ZipFile.OpenRead(refreshedOutput);
            var comments = ReadXml(archive, "word/comments.xml");
            Assert.Single(comments.Root!.Elements(W + "comment"));
            Assert.Contains("Existing note", string.Concat(comments.Descendants(W + "t").Select(element => element.Value)));
            Assert.DoesNotContain("NKU-TOOL-ANNOTATION", string.Concat(comments.Descendants(W + "t").Select(element => element.Value)));

            var document = ReadXml(archive, "word/document.xml");
            Assert.Empty(document.Descendants(W + "commentRangeStart"));
            Assert.Empty(document.Descendants(W + "commentRangeEnd"));
            Assert.Empty(document.Descendants(W + "commentReference"));
            var run = Assert.Single(document.Descendants(W + "r"), element => element.Descendants(W + "t").Any());
            Assert.Equal("pink", (string?)run.Element(W + "rPr")?.Element(W + "highlight")?.Attribute(W + "val"));
        }
        finally
        {
            File.Delete(inputPath);
            if (File.Exists(firstOutput)) File.Delete(firstOutput);
            if (File.Exists(refreshedOutput)) File.Delete(refreshedOutput);
        }
    }

    [Fact]
    public void UnsupportedLocationFailsWithoutCreatingOutput()
    {
        var inputPath = CreatePackage();
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            var diagnostic = new Diagnostic("demo/font-size", "warning", "word/header1.xml#paragraph[0]/run[0]", "12pt", "14pt", "auto", "请调整字号。");
            Assert.Throws<NotSupportedException>(() => WordAnnotationWriter.WriteAnnotatedCopy(inputPath, outputPath, [diagnostic]));
            Assert.False(File.Exists(outputPath));
        }
        finally
        {
            File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void ExistingOutputIsNotOverwritten()
    {
        var inputPath = CreatePackage();
        var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        File.WriteAllText(outputPath, "keep existing output");
        try
        {
            var diagnostic = new Diagnostic("demo/font-size", "warning", "word/document.xml#paragraph[0]/run[0]", "12pt", "14pt", "auto", "请调整字号。");
            Assert.Throws<IOException>(() => WordAnnotationWriter.WriteAnnotatedCopy(inputPath, outputPath, [diagnostic]));
            Assert.Equal("keep existing output", File.ReadAllText(outputPath));
        }
        finally
        {
            File.Delete(inputPath);
            File.Delete(outputPath);
        }
    }

    private static string CreatePackage()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
              <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml" />
            </Types>
            """);
        Add(archive, "_rels/.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" />
            </Relationships>
            """);
        Add(archive, "word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" /><w:sz w:val="28" /><w:highlight w:val="pink" /><w:u w:val="single" /></w:rPr><w:t>原文内容</w:t></w:r></w:p><w:sectPr /></w:body>
            </w:document>
            """);
        Add(archive, "word/_rels/document.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
              <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml" />
            </Relationships>
            """);
        Add(archive, "word/comments.xml", """
            <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:comment w:id="0" w:author="Author"><w:p><w:r><w:t>Existing note</w:t></w:r></w:p></w:comment>
            </w:comments>
            """);
        Add(archive, "custom/data.bin", "preserve-me");
        return path;
    }

    private static XDocument ReadXml(ZipArchive archive, string partName)
    {
        using var stream = archive.GetEntry(partName)!.Open();
        return XDocument.Load(stream);
    }

    private static string ReadEntry(ZipArchive archive, string partName)
    {
        using var reader = new StreamReader(archive.GetEntry(partName)!.Open());
        return reader.ReadToEnd();
    }

    private static string[] ReadTextRuns(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var document = ReadXml(archive, "word/document.xml");
        return document.Descendants(W + "t").Select(element => element.Value).ToArray();
    }

    private static DirectoryInfo FindRepositoryRoot()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        return directory ?? throw new DirectoryNotFoundException("Could not locate the repository root.");
    }

    private static void Add(ZipArchive archive, string name, string content)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open());
        writer.Write(content);
    }
}


