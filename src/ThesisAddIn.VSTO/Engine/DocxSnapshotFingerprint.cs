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

        private sealed class SemanticToken
        {
            public string Kind { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
            public string Attributes { get; set; } = string.Empty;
            public string Properties { get; set; } = string.Empty;
            public string Content { get; set; } = string.Empty;
        }
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
                    string[] firstOnly = a.Keys.Except(b.Keys, StringComparer.Ordinal).Take(8).ToArray();
                    string[] secondOnly = b.Keys.Except(a.Keys, StringComparer.Ordinal).Take(8).ToArray();
                    string[] changed = a.Keys.Intersect(b.Keys, StringComparer.Ordinal)
                        .Where(key => a[key] != b[key]).Take(8).ToArray();
                    return name + ": 仅快照有 [" + string.Join(", ", firstOnly)
                        + "]；仅当前文档有 [" + string.Join(", ", secondOnly)
                        + "]；值有变化 [" + string.Join(", ", changed) + "]";
                }));
            }
        }

        private static bool IsXml(string name)
        {
            return name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
        }

        private static Dictionary<string, string> ReadXmlFields(ZipArchiveEntry entry)
        {
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreWhitespace = true,
                MaxCharactersInDocument = 16 * 1024 * 1024
            };
            using (Stream stream = entry.Open())
            using (XmlReader reader = XmlReader.Create(stream, settings))
            {
                // Word can rewrite indentation between XML elements when exporting Flat OPC.
                // Preserve meaningful text (including xml:space="preserve" text in w:t),
                // but do not fingerprint serialization-only whitespace between elements.
                var root = XDocument.Load(reader).Root;
                if (root == null) throw new InvalidDataException("快照 XML 部件缺少根元素：" + entry.FullName);
                if (string.Equals(entry.FullName,"docProps/app.xml",StringComparison.OrdinalIgnoreCase)) {
                    // Word updates these derived statistics asynchronously after Save As.
                    // Keep all other application properties (for example Company) protected.
                    XNamespace ep="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
                    foreach (string name in new[] { "TotalTime", "Pages", "Words", "Characters", "Lines", "Paragraphs", "CharactersWithSpaces" })
                        root.Elements(ep+name).Remove();
                }
                if (entry.FullName.StartsWith("word/", StringComparison.Ordinal)) NormalizeExportMetadata(root, entry.FullName);
                if (string.Equals(entry.FullName, "word/comments.xml", StringComparison.OrdinalIgnoreCase))
                    NormalizeOwnedCommentPresentation(root);
                if (string.Equals(entry.FullName, "word/document.xml", StringComparison.OrdinalIgnoreCase))
                {
                    NormalizeWordAnnotations(root);
                    return ReadSemanticWordDocument(root);
                }
                var fields = new Dictionary<string, string>(StringComparer.Ordinal);
                CollectXmlFields(root, "/" + root.Name.LocalName, fields);
                return fields;
            }
        }

        private static void NormalizeOwnedCommentPresentation(XElement root)
        {
            foreach (XElement comment in root.Elements(W + "comment"))
            {
                string text = string.Join("\r", comment.Elements(W + "p").Select(paragraph =>
                    string.Concat(paragraph.Descendants().Where(element => element.Name == W + "t"
                        || element.Name == W + "tab" || element.Name == W + "br")
                        .Select(element => element.Name == W + "t" ? element.Value
                            : element.Name == W + "tab" ? "\t" : "\n"))));
                if (text.IndexOf("[NKU-TOOL-ANNOTATION:v1]", StringComparison.Ordinal) < 0) continue;
                // Word's background proofing splits styled comment runs and rewrites
                // hidden flags. The tool owns this presentation; retain every character
                // (including private location metadata) and comment identity instead.
                // Human comments retain their full formatting fingerprint.
                comment.ReplaceNodes(new XElement("nkuOwnedCommentText", text));
            }
        }

        private static Dictionary<string, string> ReadSemanticWordDocument(XElement root)
        {
            // Word can move comment/bookmark anchors and split a run when it refreshes
            // a review marker. Path-indexed XML fields then make the same text look
            // different (for example run[14] becomes run[15]). Compare the ordered
            // paragraph/format/text stream instead, while preserving meaningful XML
            // structure and every non-annotation run property.
            var tokens = new List<SemanticToken>();
            AppendSemanticTokens(root, tokens);
            var fields = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int index = 0; index < tokens.Count; index++)
            {
                SemanticToken token = tokens[index];
                fields.Add("/word/document.xml#token[" + index.ToString("D8", System.Globalization.CultureInfo.InvariantCulture) + "]",
                    string.Join("\0", new[] { token.Kind, token.Name, token.Attributes, token.Properties, token.Content }));
            }
            return fields;
        }

        private static void AppendSemanticTokens(XElement element, List<SemanticToken> tokens)
        {
            if (element.Name == W + "r")
            {
                string propertySignature = element.Element(W + "rPr") is XElement properties
                    ? CanonicalElement(properties)
                    : string.Empty;
                string content = string.Concat(element.Nodes()
                    .Where(node => !(node is XElement child && child.Name == W + "rPr"))
                    .Select(node => node is XElement text && text.Name == W + "t"
                        ? text.Value
                        : CanonicalNode(node)));
                // Word leaves empty CommentReference runs after their range markers
                // are removed. Such a run has no authored text or visible formatting.
                if (content.Length == 0) return;

                string attributes = CanonicalAttributes(element);
                if (tokens.Count > 0 && tokens[tokens.Count - 1].Kind == "run"
                    && tokens[tokens.Count - 1].Attributes == attributes
                    && tokens[tokens.Count - 1].Properties == propertySignature)
                {
                    tokens[tokens.Count - 1].Content += content;
                }
                else
                {
                    tokens.Add(new SemanticToken
                    {
                        Kind = "run",
                        Attributes = attributes,
                        Properties = propertySignature,
                        Content = content
                    });
                }
                return;
            }

            tokens.Add(new SemanticToken
            {
                Kind = "open",
                Name = element.Name.ToString(),
                Attributes = CanonicalAttributes(element)
            });
            foreach (XNode node in element.Nodes())
            {
                if (node is XElement child) AppendSemanticTokens(child, tokens);
                else if (node is XText text && !string.IsNullOrWhiteSpace(text.Value))
                    tokens.Add(new SemanticToken { Kind = "text", Content = text.Value });
            }
            tokens.Add(new SemanticToken { Kind = "close", Name = element.Name.ToString() });
        }

        private static string CanonicalNode(XNode node)
        {
            if (node is XElement element) return CanonicalElement(element);
            if (node is XText text) return "T" + text.Value.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + text.Value;
            return node.ToString(SaveOptions.DisableFormatting);
        }

        private static string CanonicalElement(XElement element)
        {
            if (element.Name == W + "lang") return string.Empty;

            var builder = new StringBuilder();
            builder.Append("E").Append(element.Name).Append('[').Append(CanonicalAttributes(element)).Append(']');
            foreach (XNode node in element.Nodes())
            {
                if (node is XElement child) builder.Append(CanonicalElement(child));
                else if (node is XText text)
                    builder.Append("T").Append(text.Value.Length.ToString(System.Globalization.CultureInfo.InvariantCulture))
                        .Append(':').Append(text.Value);
                else builder.Append(node.ToString(SaveOptions.DisableFormatting));
            }
            builder.Append("/E");
            return builder.ToString();
        }

        private static string CanonicalAttributes(XElement element)
        {
            return string.Join(";", element.Attributes()
                .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal)
                .Select(attribute => attribute.Name + "=" + attribute.Value.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    + ":" + attribute.Value));
        }

        private static void NormalizeExportMetadata(XElement root, string partName)
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

            // Word's proofing engine can add/remove range markers during background spell
            // checks. They do not alter paragraph text, formatting, or locations.
            root.DescendantsAndSelf(W + "proofErr").Remove();

            // The spelling/grammar state is refreshed asynchronously by Word itself.
            // It is not authored document content and must not invalidate an analysis.
            if (string.Equals(partName, "word/settings.xml", StringComparison.OrdinalIgnoreCase))
            {
                root.Elements(W + "proofState").Remove();
            }

            // Word may inject run/paragraph proofing-language tags while exporting its
            // in-memory document. They do not change rendered thesis formatting; removing
            // them also prevents their insertion from shifting otherwise identical text paths.
            root.DescendantsAndSelf(W + "lang").Remove();
            root.DescendantsAndSelf(W + "rPr")
                .Where(element => !element.HasElements && string.IsNullOrWhiteSpace(element.Value)).Remove();
            root.DescendantsAndSelf(W + "pPr")
                .Where(element => !element.HasElements && string.IsNullOrWhiteSpace(element.Value)).Remove();
        }

        private static void NormalizeWordAnnotations(XElement root)
        {
            // Comments and bookmarks are navigation metadata; the add-in uses both to
            // connect a highlighted source span to its issue card. They are not authored
            // thesis text or visual formatting and may be reserialized by Word on focus.
            RemoveToolGeneratedHighlights(root);
            root.DescendantsAndSelf(W + "bookmarkStart").Remove();
            root.DescendantsAndSelf(W + "bookmarkEnd").Remove();
            root.DescendantsAndSelf(W + "commentRangeStart").Remove();
            root.DescendantsAndSelf(W + "commentRangeEnd").Remove();
            root.DescendantsAndSelf(W + "commentReference").Remove();

            root.DescendantsAndSelf(W + "rPr")
                .Where(element => !element.HasElements && string.IsNullOrWhiteSpace(element.Value)).Remove();
            root.DescendantsAndSelf(W + "pPr")
                .Where(element => !element.HasElements && string.IsNullOrWhiteSpace(element.Value)).Remove();

            // Applying a background marker can split a Word run into adjacent fragments.
            // Recombine only fragments with identical direct formatting and attributes;
            // real font, size, style, or text changes remain visible to the fingerprint.
            foreach (XElement parent in root.DescendantsAndSelf().Where(element => element.HasElements).ToArray())
            {
                XNode[] nodes = parent.Nodes().ToArray();
                for (int index = 0; index + 1 < nodes.Length;)
                {
                    if (nodes[index] is XElement left && nodes[index + 1] is XElement right
                        && left.Name == W + "r" && right.Name == W + "r"
                        && HaveSameRunFormatting(left, right))
                    {
                        foreach (XNode node in new XElement(right).Nodes())
                        {
                            if (node is XElement element && element.Name == W + "rPr") continue;
                            left.Add(CloneNode(node));
                        }
                        right.Remove();
                        nodes = parent.Nodes().ToArray();
                    }
                    else
                    {
                        index++;
                    }
                }
            }

            foreach (XElement run in root.DescendantsAndSelf(W + "r"))
            {
                XNode[] nodes = run.Nodes().ToArray();
                for (int index = 0; index + 1 < nodes.Length;)
                {
                    if (nodes[index] is XElement leftText && nodes[index + 1] is XElement rightText
                        && leftText.Name == W + "t" && rightText.Name == W + "t"
                        && HaveSameAttributes(leftText, rightText))
                    {
                        leftText.Value += rightText.Value;
                        rightText.Remove();
                        nodes = run.Nodes().ToArray();
                    }
                    else
                    {
                        index++;
                    }
                }
            }
        }

        private static void RemoveToolGeneratedHighlights(XElement root)
        {
            const string bookmarkPrefix = "_NKU_MARK_";
            foreach (XElement paragraph in root.DescendantsAndSelf(W + "p"))
            {
                var starts = paragraph.DescendantsAndSelf(W + "bookmarkStart")
                    .Where(element => (element.Attribute(W + "name")?.Value ?? string.Empty)
                        .StartsWith(bookmarkPrefix, StringComparison.Ordinal))
                    .Select(element => element.Attribute(W + "id")?.Value)
                    .OfType<string>();
                var ends = new HashSet<string>(paragraph.DescendantsAndSelf(W + "bookmarkEnd")
                    .Select(element => element.Attribute(W + "id")?.Value)
                    .OfType<string>(), StringComparer.Ordinal);
                var pairedIds = new HashSet<string>(starts.Where(ends.Contains), StringComparer.Ordinal);
                if (pairedIds.Count == 0) continue;

                var activeIds = new HashSet<string>(StringComparer.Ordinal);
                foreach (XElement element in paragraph.DescendantsAndSelf())
                {
                    if (element.Name == W + "bookmarkStart")
                    {
                        string id = element.Attribute(W + "id")?.Value ?? string.Empty;
                        if (pairedIds.Contains(id)) activeIds.Add(id);
                    }
                    else if (element.Name == W + "r" && activeIds.Count > 0)
                    {
                        foreach (XElement highlight in element.Descendants(W + "highlight")
                            .Where(value => string.Equals(value.Attribute(W + "val")?.Value,
                                "yellow", StringComparison.OrdinalIgnoreCase)).ToArray())
                            highlight.Remove();
                    }
                    else if (element.Name == W + "bookmarkEnd")
                    {
                        string id = element.Attribute(W + "id")?.Value ?? string.Empty;
                        activeIds.Remove(id);
                    }
                }
            }
        }

        private static bool HaveSameRunFormatting(XElement left, XElement right)
        {
            if (!HaveSameAttributes(left, right)) return false;

            var leftProperties = left.Element(W + "rPr");
            var rightProperties = right.Element(W + "rPr");
            return (leftProperties == null && rightProperties == null)
                || (leftProperties != null && rightProperties != null && XNode.DeepEquals(leftProperties, rightProperties));
        }

        private static bool HaveSameAttributes(XElement left, XElement right)
        {
            XAttribute[] leftAttributes = left.Attributes().OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).ToArray();
            XAttribute[] rightAttributes = right.Attributes().OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).ToArray();
            if (leftAttributes.Length != rightAttributes.Length) return false;
            for (int index = 0; index < leftAttributes.Length; index++)
            {
                if (leftAttributes[index].Name != rightAttributes[index].Name
                    || leftAttributes[index].Value != rightAttributes[index].Value) return false;
            }
            return true;
        }

        private static XNode CloneNode(XNode node)
        {
            if (node is XElement element) return new XElement(element);
            if (node is XText text) return new XText(text.Value);
            if (node is XComment comment) return new XComment(comment.Value);
            if (node is XProcessingInstruction instruction) return new XProcessingInstruction(instruction.Target, instruction.Data);
            if (node is XDocumentType documentType)
                return new XDocumentType(documentType.Name, documentType.PublicId, documentType.SystemId, documentType.InternalSubset);
            throw new InvalidDataException("快照含无法安全复制的 XML 节点类型：" + node.NodeType);
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

