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
    /// <param name="inputPath">输入 docx 路径</param>
    /// <param name="outputPath">输出 docx 路径</param>
    /// <param name="styleId">样式 ID（如 Heading1、Normal）</param>
    /// <param name="fontSizePt">字号（磅，如三号=16、小四=12）</param>
    public void SetStyleFontSize(string inputPath, string outputPath, string styleId, double fontSizePt)
    {
        // Word 里字号单位是 half-points（半磅），所以要乘 2
        var halfPoints = (int)(fontSizePt * 2);

        using var inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(inputStream, ZipArchiveMode.Read);

        // 读取 styles.xml
        var stylesEntry = archive.GetEntry("word/styles.xml")
            ?? throw new InvalidDataException("文档中没有 styles.xml");

        XDocument stylesDoc;
        using (var reader = new StreamReader(stylesEntry.Open()))
        {
            stylesDoc = XDocument.Load(reader);
        }

        // 找到目标样式
        var styleElement = stylesDoc.Root?
            .Elements(W + "style")
            .FirstOrDefault(s => s.Attribute(W + "styleId")?.Value == styleId)
            ?? throw new ArgumentException($"找不到样式：{styleId}");

        // 找到或创建 rPr
        var rPr = styleElement.Element(W + "rPr");
        if (rPr == null)
        {
            rPr = new XElement(W + "rPr");
            styleElement.Add(rPr);
        }

        // 设置字号
        var szElement = rPr.Element(W + "sz");
        if (szElement == null)
        {
            szElement = new XElement(W + "sz");
            rPr.Add(szElement);
        }
        szElement.SetAttributeValue(W + "val", halfPoints.ToString());

        // 同时设置 complex script 字号（中文/复杂文种）
        var szCsElement = rPr.Element(W + "szCs");
        if (szCsElement == null)
        {
            szCsElement = new XElement(W + "szCs");
            rPr.Add(szCsElement);
        }
        szCsElement.SetAttributeValue(W + "val", halfPoints.ToString());

        // 保存新的 docx
        File.Copy(inputPath, outputPath, overwrite: true);

        using var outputStream = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite);
        using var outputArchive = new ZipArchive(outputStream, ZipArchiveMode.Update);

        // 删除旧的 styles.xml
        outputArchive.GetEntry("word/styles.xml")?.Delete();

        // 写入新的 styles.xml
        var newEntry = outputArchive.CreateEntry("word/styles.xml");
        using (var writer = new StreamWriter(newEntry.Open()))
        {
            stylesDoc.Save(writer);
        }
    }

    /// <summary>
    /// 修改指定样式的字体名称。
    /// </summary>
    public void SetStyleFontName(string inputPath, string outputPath, string styleId, string eastAsiaFont, string asciiFont)
    {
        using var inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(inputStream, ZipArchiveMode.Read);

        var stylesEntry = archive.GetEntry("word/styles.xml")
            ?? throw new InvalidDataException("文档中没有 styles.xml");

        XDocument stylesDoc;
        using (var reader = new StreamReader(stylesEntry.Open()))
        {
            stylesDoc = XDocument.Load(reader);
        }

        var styleElement = stylesDoc.Root?
            .Elements(W + "style")
            .FirstOrDefault(s => s.Attribute(W + "styleId")?.Value == styleId)
            ?? throw new ArgumentException($"找不到样式：{styleId}");

        var rPr = styleElement.Element(W + "rPr");
        if (rPr == null)
        {
            rPr = new XElement(W + "rPr");
            styleElement.Add(rPr);
        }

        // 设置字体
        var rFonts = rPr.Element(W + "rFonts");
        if (rFonts == null)
        {
            rFonts = new XElement(W + "rFonts");
            rPr.Add(rFonts);
        }
        rFonts.SetAttributeValue(W + "eastAsia", eastAsiaFont);
        rFonts.SetAttributeValue(W + "ascii", asciiFont);
        rFonts.SetAttributeValue(W + "hAnsi", asciiFont);

        File.Copy(inputPath, outputPath, overwrite: true);

        using var outputStream = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite);
        using var outputArchive = new ZipArchive(outputStream, ZipArchiveMode.Update);

        outputArchive.GetEntry("word/styles.xml")?.Delete();

        var newEntry = outputArchive.CreateEntry("word/styles.xml");
        using (var writer = new StreamWriter(newEntry.Open()))
        {
            stylesDoc.Save(writer);
        }
    }
}
