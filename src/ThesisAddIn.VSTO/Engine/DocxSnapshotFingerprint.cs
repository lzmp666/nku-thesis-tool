using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ThesisAddIn.VSTO.Engine
{
    /// <summary>Compares package content while ignoring ZIP and XML serialization metadata.</summary>
    public static class DocxSnapshotFingerprint
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        private static readonly HashSet<string> RevisionIds = new HashSet<string>(StringComparer.Ordinal)
            { "rsidR", "rsidRDefault", "rsidRPr", "rsidP", "rsidDel", "rsidSect" };
        /// <summary>Reports only package part names, never document text, for mismatch diagnostics.</summary>
        public static string[] GetChangedParts(string firstPath, string secondPath)
        {
            Dictionary<string, byte[]> first = ReadPartHashes(firstPath);
            Dictionary<string, byte[]> second = ReadPartHashes(secondPath);
            return first.Keys.Union(second.Keys, StringComparer.Ordinal).Where(name =>
            {
                return !first.ContainsKey(name) || !second.ContainsKey(name) || !first[name].SequenceEqual(second[name]);
            }).OrderBy(name => name, StringComparer.Ordinal).ToArray();
        }

        private static Dictionary<string, byte[]> ReadPartHashes(string path)
        {
            var hashes = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            using (var file = File.OpenRead(path))
            using (var archive = new ZipArchive(file, ZipArchiveMode.Read))
            using (var hash = SHA256.Create())
            {
                foreach (ZipArchiveEntry entry in archive.Entries.Where(item =>
                    !string.Equals(item.FullName, "docProps/core.xml", StringComparison.OrdinalIgnoreCase)))
                {
                    if (IsXml(entry.FullName)) hashes.Add(entry.FullName, HashXmlFields(ReadXmlFields(entry)));
                    else using (Stream content = entry.Open()) hashes.Add(entry.FullName, hash.ComputeHash(content));
                }
            }
            return hashes;
        }

        public static string DescribeDifferences(string firstPath, string secondPath)
        {
            string[] names = GetChangedParts(firstPath, secondPath);
            using (var firstFile = File.OpenRead(firstPath))
            using (var secondFile = File.OpenRead(secondPath))
            using (var first = new ZipArchive(firstFile, ZipArchiveMode.Read))
            using (var second = new ZipArchive(secondFile, ZipArchiveMode.Read))
            {
                return string.Join("; ", names.Select(name =>
                {
                    var left = first.GetEntry(name);
                    var right = second.GetEntry(name);
                    if (!IsXml(name) || left == null || right == null) return name;
                    Dictionary<string, string> a = ReadXmlFields(left);
                    Dictionary<string, string> b = ReadXmlFields(right);
                    string[] fields = a.Keys.Union(b.Keys, StringComparer.Ordinal).Where(key =>
                        !a.ContainsKey(key) || !b.ContainsKey(key) || a[key] != b[key]).Take(12).ToArray();
                    return name + ": " + string.Join(", ", fields);
                }));
            }
        }

        private static bool IsXml(string name)
        {
            return name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
        }

        private static Dictionary<string, string> ReadXmlFields(ZipArchiveEntry entry)
        {
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 16 * 1024 * 1024 };
            using (Stream stream = entry.Open())
            using (XmlReader reader = XmlReader.Create(stream, settings))
            {
                var root = XDocument.Load(reader, LoadOptions.PreserveWhitespace).Root;
                if (root == null) throw new InvalidDataException("快照 XML 部件缺少根元素：" + entry.FullName);
                if (entry.FullName.StartsWith("word/", StringComparison.Ordinal)) NormalizeExportIds(root);
                var fields = new Dictionary<string, string>(StringComparer.Ordinal);
                CollectXmlFields(root, "/" + root.Name.LocalName, fields);
                return fields;
            }
        }

        private static void NormalizeExportIds(XElement root)
        {
            // RSIDs identify editing sessions, not text/format values (ECMA-376 17.3.1.22):
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.paragraph
            foreach (XAttribute attribute in root.DescendantsAndSelf().Attributes().Where(a =>
                a.Name.Namespace == W && RevisionIds.Contains(a.Name.LocalName)).ToArray()) attribute.Remove();
            if (root.Name == W + "settings") root.Elements(W + "rsids").Remove();

            // WordOpenXML regenerates paragraph IDs for automatic note separators.
            // Retain IDs for body paragraphs and ordinary notes: they can anchor comments.
            foreach (XElement paragraph in root.DescendantsAndSelf(W + "p"))
            {
                bool separator = paragraph.Ancestors().Any(note =>
                    (note.Name == W + "footnote" || note.Name == W + "endnote") &&
                    (note.Attribute(W + "type")?.Value == "separator" || note.Attribute(W + "type")?.Value == "continuationSeparator"));
                if (separator) paragraph.Attribute(W14 + "paraId")?.Remove();
            }
        }

        private static void CollectXmlFields(XElement element, string path, Dictionary<string, string> fields)
        {
            fields.Add(path + "#name", element.Name.ToString());
            foreach (XAttribute attribute in element.Attributes())
                fields.Add(path + "@" + attribute.Name, attribute.Value);
            int index = 0;
            foreach (XNode node in element.Nodes())
            {
                var child = node as XElement;
                if (child != null) CollectXmlFields(child, path + "/" + child.Name.LocalName + "[" + index + "]", fields);
                else fields.Add(path + "/node[" + index + "]", node.ToString(SaveOptions.DisableFormatting));
                index++;
            }
        }

        private static byte[] HashXmlFields(Dictionary<string, string> fields)
        {
            using (var bytes = new MemoryStream())
            using (var writer = new BinaryWriter(bytes, new UTF8Encoding(false), leaveOpen: true))
            using (var hash = SHA256.Create())
            {
                foreach (var field in fields.OrderBy(item => item.Key, StringComparer.Ordinal))
                {
                    writer.Write(field.Key);
                    writer.Write(field.Value);
                }
                writer.Flush();
                bytes.Position = 0;
                return hash.ComputeHash(bytes);
            }
        }

        public static string Compute(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("必须提供 DOCX 文件路径。", nameof(path));
            }

            using (var fingerprint = IncrementalHash.CreateHash(HashAlgorithmName.SHA256))
            {
                foreach (var entry in ReadPartHashes(path).OrderBy(item => item.Key, StringComparer.Ordinal))
                {
                    fingerprint.AppendData(Encoding.UTF8.GetBytes(entry.Key));
                    fingerprint.AppendData(new byte[] { 0 });
                    fingerprint.AppendData(entry.Value);
                }

                return BitConverter.ToString(fingerprint.GetHashAndReset()).Replace("-", string.Empty).ToLowerInvariant();
            }
        }
    }
}

