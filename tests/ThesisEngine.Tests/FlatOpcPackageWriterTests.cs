using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using ThesisAddIn.VSTO.Engine;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class FlatOpcPackageWriterTests
{
    private const string PackageNamespace = "http://schemas.microsoft.com/office/2006/xmlPackage";
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    [Fact]
    public void WritesXmlBinaryAndRelationshipPartsToDocxPackage()
    {
        var path = Path.Combine(Path.GetTempPath(), $"flat-opc-{Guid.NewGuid():N}.docx");
        try
        {
            FlatOpcPackageWriter.Write(CreateFlatOpc("未保存的内存内容"), path);

            using var archive = ZipFile.OpenRead(path);
            var names = archive.Entries.Select(entry => entry.FullName).ToArray();
            Assert.Contains("[Content_Types].xml", names);
            Assert.Contains("_rels/.rels", names);
            Assert.Contains("word/document.xml", names);
            Assert.Contains("word/media/image1.png", names);

            var document = ReadXml(archive, "word/document.xml");
            Assert.Contains("未保存的内存内容", document);
            Assert.Equal(new byte[] { 1, 2, 3, 4 }, ReadBytes(archive, "word/media/image1.png"));

            var contentTypes = XDocument.Parse(ReadXml(archive, "[Content_Types].xml"));
            Assert.Contains(contentTypes.Root!.Elements(XName.Get("Default", ContentTypesNamespace)),
                element => (string?)element.Attribute("Extension") == "rels"
                    && (string?)element.Attribute("ContentType") == "application/vnd.openxmlformats-package.relationships+xml");
            var documentOverride = contentTypes.Root!
                .Elements(XName.Get("Override", ContentTypesNamespace))
                .Single(element => (string?)element.Attribute("PartName") == "/word/document.xml");
            Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                (string?)documentOverride.Attribute("ContentType"));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void RejectsSnapshotWithoutDocumentXmlAndDoesNotCreateOutput()
    {
        var path = Path.Combine(Path.GetTempPath(), $"flat-opc-invalid-{Guid.NewGuid():N}.docx");
        const string invalid = "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/word/styles.xml\" pkg:contentType=\"application/xml\"><pkg:xmlData><styles xmlns=\"urn:test\"/></pkg:xmlData></pkg:part></pkg:package>";
        try
        {
            Assert.Throws<InvalidDataException>(() => FlatOpcPackageWriter.Write(invalid, path));
            Assert.False(File.Exists(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void RefusesToOverwriteAnExistingSnapshot()
    {
        var path = Path.Combine(Path.GetTempPath(), $"flat-opc-existing-{Guid.NewGuid():N}.docx");
        File.WriteAllText(path, "keep");
        try
        {
            Assert.Throws<IOException>(() => FlatOpcPackageWriter.Write(CreateFlatOpc("new"), path));
            Assert.Equal("keep", File.ReadAllText(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static string CreateFlatOpc(string text)
    {
        return $"""
            <pkg:package xmlns:pkg="{PackageNamespace}">
              <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
                <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" /></Relationships></pkg:xmlData>
              </pkg:part>
              <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
                <pkg:xmlData><document xmlns="urn:test"><body><p>{System.Security.SecurityElement.Escape(text)}</p></body></document></pkg:xmlData>
              </pkg:part>
              <pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png" pkg:compression="store">
                <pkg:binaryData>AQIDBA==</pkg:binaryData>
              </pkg:part>
            </pkg:package>
            """;
    }

    private static string ReadXml(ZipArchive archive, string entryName)
    {
        using var stream = archive.GetEntry(entryName)!.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static byte[] ReadBytes(ZipArchive archive, string entryName)
    {
        using var stream = archive.GetEntry(entryName)!.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }
}

