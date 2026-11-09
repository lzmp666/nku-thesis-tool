using System.IO.Compression;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// Word 样式重写器。
/// 修改样式表中的样式定义，所有套用该样式的段落会自动更新。
/// </summary>
public sealed class StyleRewriter
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = MainNamespace;

    /// <summary>
    /// 修改指定样式的字号。
    /// </summary>
    public void SetStyleFontSize(string inputPath, string outputPath, string styleId, double fontSizePt)
    {
        var halfPoints = (int)(fontSizePt * 2);

        // 第一步：读 styles.xml
        var stylesDoc = ReadStylesXml(inputPath);

        // 修改 stylesDoc
        var styleElement = FindStyle(stylesDoc, styleId)
            ?? throw new ArgumentException($"找不到样式：{styleId}");

        var rPr = styleElement.Element(W + "rPr");
        if (rPr == null)
        {
            rPr = new XElement(W + "rPr");
            styleElement.Add(rPr);
        }

        var szElement = rPr.Element(W + "sz");
        if (szElement == null)
        {
            szElement = new XElement(W + "sz");
            rPr.Add(szElement);
        }
        szElement.SetAttributeValue(W + "val", halfPoints.ToString());

        var szCsElement = rPr.Element(W + "szCs");
        if (szCsElement == null)
        {
            szCsElement = new XElement(W + "szCs");
            rPr.Add(szCsElement);
        }
        szCsElement.SetAttributeValue(W + "val", halfPoints.ToString());

        // 第二步：写入 styles.xml
        WriteStylesXml(inputPath, outputPath, stylesDoc);
    }

    /// <summary>
    /// 修改指定样式的字体名称。
    /// </summary>
    public void SetStyleFontName(string inputPath, string outputPath, string styleId, string eastAsiaFont, string asciiFont)
    {
        // 第一步：读 styles.xml
        var stylesDoc = ReadStylesXml(inputPath);

        // 修改 stylesDoc
        var styleElement = FindStyle(stylesDoc, styleId)
            ?? throw new ArgumentException($"找不到样式：{styleId}");

        var rPr = styleElement.Element(W + "rPr");
        if (rPr == null)
        {
            rPr = new XElement(W + "rPr");
            styleElement.Add(rPr);
        }

        var rFonts = rPr.Element(W + "rFonts");
        if (rFonts == null)
        {
            rFonts = new XElement(W + "rFonts");
            rPr.Add(rFonts);
        }
        rFonts.SetAttributeValue(W + "eastAsia", eastAsiaFont);
        rFonts.SetAttributeValue(W + "ascii", asciiFont);
        rFonts.SetAttributeValue(W + "hAnsi", asciiFont);

        // 第二步：写入 styles.xml
        WriteStylesXml(inputPath, outputPath, stylesDoc);
    }

    /// <summary>
    /// 修改指定样式的行距。
    /// </summary>
    /// <param name="lineSpacingMultiplier">行距倍数：1.0 = 单倍，1.5 = 1.5倍，2.0 = 双倍</param>
    public void SetStyleLineSpacing(string inputPath, string outputPath, string styleId, double lineSpacingMultiplier)
    {
        // Word 行距单位是 1/240 行
        var lineValue = (int)(lineSpacingMultiplier * 240);

        var stylesDoc = ReadStylesXml(inputPath);

        var styleElement = FindStyle(stylesDoc, styleId)
            ?? throw new ArgumentException($"找不到样式：{styleId}");

        var pPr = styleElement.Element(W + "pPr");
        if (pPr == null)
        {
            pPr = new XElement(W + "pPr");
            styleElement.Add(pPr);
        }

        var spacing = pPr.Element(W + "spacing");
        if (spacing == null)
        {
            spacing = new XElement(W + "spacing");
            pPr.Add(spacing);
        }
        spacing.SetAttributeValue(W + "line", lineValue.ToString());
        spacing.SetAttributeValue(W + "lineRule", "auto");

        WriteStylesXml(inputPath, outputPath, stylesDoc);
    }

    /// <summary>
    /// 文档里是否存在可用的样式（风格 id 或内置名能对上其一即可）。
    /// </summary>
    public bool HasStyle(string docxPath, string styleId)
    {
        var stylesDoc = ReadStylesXml(docxPath);
        return FindStyle(stylesDoc, styleId) != null;
    }

    /// <summary>
    /// 按 styleId 找样式，找不到再按内置名兜底。
    /// </summary>
    /// <remarks>
    /// styleId 不是规范属性，只是文档内部的名字。Word 中文版、WPS、以及部分
    /// 第三方库重写过的文档，常把正文样式写成 <c>w:styleId="a"</c>，但
    /// <c>w:name</c> 仍然是 <c>Normal</c>。只按 styleId 精确匹配会直接判定
    /// 「文档没有正文样式」，整条格式化链在这里断掉。
    /// </remarks>
    private static XElement? FindStyle(XDocument stylesDoc, string styleId)
    {
        var styles = stylesDoc.Root?.Elements(W + "style").ToList();
        if (styles == null || styles.Count == 0) return null;

        var byId = styles.FirstOrDefault(s => s.Attribute(W + "styleId")?.Value == styleId);
        if (byId != null) return byId;

        // 内置名可能带空格（heading 1），去掉空格再比，大小写不敏感。
        var target = NormalizeStyleName(styleId);
        return styles.FirstOrDefault(s =>
            string.Equals(
                NormalizeStyleName(s.Element(W + "name")?.Attribute(W + "val")?.Value),
                target,
                StringComparison.OrdinalIgnoreCase));
    }

    private static string NormalizeStyleName(string? name)
        => name?.Replace(" ", "").Replace("\u3000", "") ?? string.Empty;

    private static XDocument ReadStylesXml(string docxPath)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var stylesEntry = archive.GetEntry("word/styles.xml")
            ?? throw new InvalidDataException("文档中没有 styles.xml");

        using var reader = new StreamReader(stylesEntry.Open());
        return XDocument.Load(reader);
    }

    private static void WriteStylesXml(string inputPath, string outputPath, XDocument stylesDoc)
    {
        // 如果输入和输出不是同一个文件，先复制
        if (!string.Equals(Path.GetFullPath(inputPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(inputPath, outputPath, overwrite: true);
        }

        using var stream = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update);

        archive.GetEntry("word/styles.xml")?.Delete();

        var newEntry = archive.CreateEntry("word/styles.xml");
        using var writer = new StreamWriter(newEntry.Open());
        stylesDoc.Save(writer);
    }
}
