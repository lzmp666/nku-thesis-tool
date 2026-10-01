using System.IO.Compression;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class DocumentLoaderTests
{
    [Fact]
    public void ValidPackageLoadsWithoutChangingInput()
    {
        var path = CreatePackage(true);
        try
        {
            var before = File.ReadAllBytes(path);
            var result = DocumentLoader.Load(path);
            Assert.Equal(3, result.PartCount);
            Assert.Equal(Path.GetFullPath(path), result.Path);
            Assert.Equal(before, File.ReadAllBytes(path)); // read-only loading, not repair idempotence
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void MissingDocumentPartIsRejected()
    {
        var path = CreatePackage(false);
        try { Assert.Throws<InvalidDataException>(() => DocumentLoader.Load(path)); }
        finally { File.Delete(path); }
    }

    [Fact]
    public void WrongExtensionIsRejected() =>
        Assert.Throws<ArgumentException>(() => DocumentLoader.Load("example.txt"));

    [Fact]
    public void CorruptArchiveIsRejected()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        File.WriteAllText(path, "not a ZIP archive");
        try { Assert.Throws<InvalidDataException>(() => DocumentLoader.Load(path)); }
        finally { File.Delete(path); }
    }

    [Fact]
    public void MissingFileIsRejected() =>
        Assert.Throws<FileNotFoundException>(() => DocumentLoader.Load(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx")));

    private static string CreatePackage(bool includeDocument)
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml", """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>""");
        Add(archive, "_rels/.rels", """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""");
        if (includeDocument) Add(archive, "word/document.xml", """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Synthetic thesis</w:t></w:r></w:p></w:body></w:document>""");
        return path;
    }

    private static void Add(ZipArchive archive, string name, string xml)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open());
        writer.Write(xml);
    }
}
