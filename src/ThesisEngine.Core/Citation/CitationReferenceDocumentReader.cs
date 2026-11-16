using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace ThesisEngine.Core.Citation;

public sealed class CitationDocumentReader
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public CitationDocument Read(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        if (!string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Only .docx files are supported.", nameof(path));

        using var archive = ZipFile.OpenRead(Path.GetFullPath(path));
        var entries = archive.Entries.Where(item => item.FullName == "word/document.xml").ToArray();
        if (entries.Length != 1)
            throw new InvalidDataException(entries.Length == 0
                ? "Missing OOXML part: word/document.xml"
                : "Duplicate OOXML part: word/document.xml");
        var entry = entries[0];
        if (entry.Length > 16 * 1024 * 1024)
            throw new InvalidDataException("OOXML part too large: word/document.xml");

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = 16 * 1024 * 1024
        });
        var document = XDocument.Load(reader);
        if (document.Root?.Name != W + "document")
            throw new InvalidDataException("Invalid OOXML root: word/document.xml");
        var body = document.Root?.Element(W + "body")
            ?? throw new InvalidDataException("Missing Word document body.");
        var paragraphs = body.Descendants(W + "p")
            .Where(paragraph => !paragraph.Ancestors(W + "txbxContent").Any())
            .Select((paragraph, index) => ReadParagraph(paragraph, index))
            .ToArray();
        var citationFields = body.Descendants(W + "instrText")
            .Count(element => IsCitationField(element.Value));

        return new CitationDocument(paragraphs, citationFields);
    }

    private static CitationParagraph ReadParagraph(XElement paragraph, int paragraphIndex)
    {
        var text = new System.Text.StringBuilder();
        var runs = new List<CitationRunSpan>();
        foreach (var run in paragraph.Descendants(W + "r").Where(run => !run.Ancestors(W + "p").Skip(1).Any()))
        {
            var runText = string.Concat(run.Descendants(W + "t").Select(element => element.Value));
            if (runText.Length == 0) continue;
            var start = text.Length;
            text.Append(runText);
            var alignment = (string?)run.Element(W + "rPr")?.Element(W + "vertAlign")?.Attribute(W + "val");
            runs.Add(new CitationRunSpan(start, runText.Length,
                string.Equals(alignment, "superscript", StringComparison.OrdinalIgnoreCase)));
        }
        return new CitationParagraph(paragraphIndex, text.ToString(), runs);
    }

    private static bool IsCitationField(string instruction) =>
        instruction.Contains("ZOTERO_ITEM", StringComparison.OrdinalIgnoreCase) ||
        instruction.Contains("CSL_CITATION", StringComparison.OrdinalIgnoreCase) ||
        instruction.Contains("ADDIN EN.CITE", StringComparison.OrdinalIgnoreCase) ||
        instruction.Contains("CITATION ", StringComparison.OrdinalIgnoreCase);
}
