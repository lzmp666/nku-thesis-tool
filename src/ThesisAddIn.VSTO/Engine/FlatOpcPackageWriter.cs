#nullable disable
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ThesisAddIn.VSTO.Engine
{
    /// <summary>
    /// Converts Word's in-memory Flat OPC representation into a standalone DOCX package.
    /// </summary>
    public static class FlatOpcPackageWriter
    {
        private static readonly XNamespace PackageNamespace = "http://schemas.microsoft.com/office/2006/xmlPackage";
        private static readonly XNamespace ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

        public static void Write(string flatOpc, string outputPath)
        {
            if (string.IsNullOrWhiteSpace(flatOpc))
            {
                throw new ArgumentException("Word returned an empty Open XML snapshot.", nameof(flatOpc));
            }
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new ArgumentException("An output path is required.", nameof(outputPath));
            }

            XDocument package = ParseFlatOpc(flatOpc);
            List<FlatOpcPart> parts = package.Root
                .Elements(PackageNamespace + "part")
                .Select(ReadPart)
                .Where(part => !string.Equals(part.Name, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (!parts.Any(part => string.Equals(part.Name, "word/document.xml", StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidDataException("Word's Open XML snapshot does not contain word/document.xml.");
            }
            if (parts.Select(part => part.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != parts.Count)
            {
                throw new InvalidDataException("Word's Open XML snapshot contains duplicate package parts.");
            }

            string fullPath = Path.GetFullPath(outputPath);
            string directory = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrEmpty(directory))
            {
                throw new InvalidDataException("The DOCX output path has no parent directory.");
            }
            Directory.CreateDirectory(directory);
            string temporaryPath = fullPath + "." + Guid.NewGuid().ToString("N") + ".tmp";

            try
            {
                using (FileStream stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, false))
                {
                    WriteContentTypes(archive, parts);
                    foreach (FlatOpcPart part in parts)
                    {
                        ZipArchiveEntry entry = archive.CreateEntry(part.Name, CompressionLevel.Optimal);
                        using (Stream entryStream = entry.Open())
                        {
                            if (part.XmlContent != null)
                            {
                                XmlWriterSettings settings = new XmlWriterSettings
                                {
                                    Encoding = new UTF8Encoding(false),
                                    OmitXmlDeclaration = false,
                                    Indent = false,
                                    CloseOutput = false
                                };
                                using (XmlWriter writer = XmlWriter.Create(entryStream, settings))
                                {
                                    part.XmlContent.WriteTo(writer);
                                }
                            }
                            else
                            {
                                entryStream.Write(part.BinaryContent, 0, part.BinaryContent.Length);
                            }
                        }
                    }
                }

                if (File.Exists(fullPath))
                {
                    throw new IOException("Refusing to overwrite an existing snapshot file.");
                }
                File.Move(temporaryPath, fullPath);
            }
            finally
            {
                if (File.Exists(temporaryPath))
                {
                    File.Delete(temporaryPath);
                }
            }
        }

        private static XDocument ParseFlatOpc(string flatOpc)
        {
            XmlReaderSettings settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = 0
            };
            using (StringReader text = new StringReader(flatOpc))
            using (XmlReader reader = XmlReader.Create(text, settings))
            {
                XDocument document = XDocument.Load(reader, LoadOptions.None);
                if (document.Root == null || document.Root.Name != PackageNamespace + "package")
                {
                    throw new InvalidDataException("Word returned an unrecognized Flat OPC package.");
                }
                return document;
            }
        }

        private static FlatOpcPart ReadPart(XElement element)
        {
            string rawName = (string)element.Attribute(PackageNamespace + "name");
            string contentType = (string)element.Attribute(PackageNamespace + "contentType");
            if (string.IsNullOrWhiteSpace(rawName) || string.IsNullOrWhiteSpace(contentType))
            {
                throw new InvalidDataException("A Flat OPC part is missing its name or content type.");
            }

            string name = Uri.UnescapeDataString(rawName.TrimStart('/'));
            if (string.IsNullOrWhiteSpace(name) || name.IndexOf('\\') >= 0
                || name.Split('/').Any(segment => segment == ".." || segment == "."))
            {
                throw new InvalidDataException("Word returned an invalid package part name.");
            }

            XElement xmlData = element.Element(PackageNamespace + "xmlData");
            if (xmlData != null)
            {
                XElement[] roots = xmlData.Elements().ToArray();
                if (roots.Length != 1)
                {
                    throw new InvalidDataException("An XML package part must contain exactly one root element.");
                }
                return new FlatOpcPart(name, contentType, new XDocument(new XDeclaration("1.0", "UTF-8", null), roots[0]), null);
            }

            XElement binaryData = element.Element(PackageNamespace + "binaryData");
            if (binaryData == null)
            {
                throw new InvalidDataException("A Flat OPC part contains neither XML nor binary data.");
            }
            return new FlatOpcPart(name, contentType, null, Convert.FromBase64String(binaryData.Value));
        }

        private static void WriteContentTypes(ZipArchive archive, IEnumerable<FlatOpcPart> parts)
        {
            XElement types = new XElement(
                ContentTypesNamespace + "Types",
                new XElement(
                    ContentTypesNamespace + "Default",
                    new XAttribute("Extension", "rels"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")));
            foreach (FlatOpcPart part in parts)
            {
                if (part.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                types.Add(new XElement(
                    ContentTypesNamespace + "Override",
                    new XAttribute("PartName", "/" + part.Name),
                    new XAttribute("ContentType", part.ContentType)));
            }

            ZipArchiveEntry entry = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
            using (Stream entryStream = entry.Open())
            using (XmlWriter writer = XmlWriter.Create(entryStream, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                OmitXmlDeclaration = false,
                Indent = false,
                CloseOutput = false
            }))
            {
                new XDocument(new XDeclaration("1.0", "UTF-8", null), types).WriteTo(writer);
            }
        }

        private sealed class FlatOpcPart
        {
            public FlatOpcPart(string name, string contentType, XDocument xmlContent, byte[] binaryContent)
            {
                Name = name;
                ContentType = contentType;
                XmlContent = xmlContent;
                BinaryContent = binaryContent;
            }

            public string Name { get; }
            public string ContentType { get; }
            public XDocument XmlContent { get; }
            public byte[] BinaryContent { get; }
        }
    }
}

