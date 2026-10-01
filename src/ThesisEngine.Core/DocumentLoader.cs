using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace ThesisEngine.Core;

public sealed record LoadedDocument(string Path, int PartCount);

public static class DocumentLoader
{
    public static LoadedDocument Load(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        if (!string.Equals(System.IO.Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Only .docx files are supported.", nameof(path));
        var fullPath = System.IO.Path.GetFullPath(path);
        using var archive = ZipFile.OpenRead(fullPath);
        ReadPart(archive, "[Content_Types].xml", "Types", "http://schemas.openxmlformats.org/package/2006/content-types");
        ReadPart(archive, "_rels/.rels", "Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
        var document = ReadPart(archive, "word/document.xml", "document", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        if (document.Root!.Element(document.Root.Name.Namespace + "body") is null)
            throw new InvalidDataException("Missing Word document body.");
        return new LoadedDocument(fullPath, archive.Entries.Count);
    }

    private static XDocument ReadPart(ZipArchive archive, string name, string root, string ns)
    {
        var matches = archive.Entries.Where(e => e.FullName == name).ToArray();
        if (matches.Length != 1) throw new InvalidDataException($"Missing or duplicate OOXML part: {name}");
        if (matches[0].Length > 16 * 1024 * 1024) throw new InvalidDataException($"OOXML part too large: {name}");
        using var stream = matches[0].Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = 16 * 1024 * 1024
        });
        var xml = XDocument.Load(reader);
        if (xml.Root?.Name != XName.Get(root, ns)) throw new InvalidDataException($"Invalid OOXML root: {name}");
        return xml;
    }
}
