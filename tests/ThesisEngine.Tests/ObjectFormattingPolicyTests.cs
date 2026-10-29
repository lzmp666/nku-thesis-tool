using System.Xml.Linq;
using ThesisAddIn.VSTO.Engine;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class ObjectFormattingPolicyTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static XElement Table() => XElement.Parse($"""
        <w:tbl xmlns:w="{W}"><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>名称</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:tc><w:p><w:r><w:t>样本</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>12.3</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
        """);

    [Fact]
    public void AcceptsRectangularBilingualTableWithoutDependingOnVisibleTextOrTableStyle()
    {
        var table = Table();
        table.Element(XName.Get("tblPr", W))!.Add(new XElement(XName.Get("tblStyle", W), new XAttribute(XName.Get("val", W), "TableGrid")));
        Assert.Null(ObjectFormattingPolicy.TableRejection(table.ToString()));
    }

    [Theory]
    [InlineData("gridSpan")]
    [InlineData("vMerge")]
    [InlineData("fldChar")]
    [InlineData("drawing")]
    [InlineData("ins")]
    [InlineData("sdt")]
    public void RejectsStructuresWhoseContentOrHeaderCannotBeSafelyReformatted(string element)
    {
        var table = Table(); table.Descendants(XName.Get("tc", W)).First().Add(new XElement(XName.Get(element, W)));
        Assert.NotNull(ObjectFormattingPolicy.TableRejection(table.ToString()));
    }

    [Fact]
    public void RejectsNestedAndIrregularTablesAndEquations()
    {
        var nested = Table(); nested.Descendants(XName.Get("tc", W)).First().Add(Table());
        Assert.NotNull(ObjectFormattingPolicy.TableRejection(nested.ToString()));
        var irregular = Table(); irregular.Descendants(XName.Get("tc", W)).Last().Remove();
        Assert.NotNull(ObjectFormattingPolicy.TableRejection(irregular.ToString()));
        var math = Table(); math.Add(new XElement(XName.Get("oMath", "http://schemas.openxmlformats.org/officeDocument/2006/math")));
        Assert.NotNull(ObjectFormattingPolicy.TableRejection(math.ToString()));
    }

    [Fact]
    public void FlatOpcGlossaryTablesDoNotMakeAnOrdinarySelectionLookNested()
    {
        XNamespace package = "http://schemas.microsoft.com/office/2006/xmlPackage";
        var xml = new XElement(package + "package",
            new XElement(package + "part", new XAttribute(package + "name", "/word/document.xml"),
                new XElement(package + "xmlData", new XElement(XName.Get("document", W), new XElement(XName.Get("body", W), Table())))),
            new XElement(package + "part", new XAttribute(package + "name", "/word/glossary/document.xml"),
                new XElement(package + "xmlData", new XElement(XName.Get("glossaryDocument", W), Table()))));
        Assert.Null(ObjectFormattingPolicy.TableRejection(xml.ToString()));
    }

    [Fact]
    public void SelectionWithPeerTablesMustMatchExactlyOneAuthoredTextStream()
    {
        var first = Table(); var other = Table(); other.Descendants(XName.Get("t", W)).First().Value = "Other";
        var body = new XElement(XName.Get("body", W), first, other);
        Assert.Null(ObjectFormattingPolicy.TableRejection(body.ToString(), "名称\r\aValue\r\a样本\r\a12.3\r\a"));
        Assert.NotNull(ObjectFormattingPolicy.TableRejection(body.ToString(), "unrelated"));
        other.Descendants(XName.Get("t", W)).First().Value = "名称";
        Assert.NotNull(ObjectFormattingPolicy.TableRejection(body.ToString(), "名称Value样本12.3"));
    }
}
