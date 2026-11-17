using System.Xml;
using System.Text;
using ThesisEngine.Cli;
using ThesisEngine.Core;

var utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
Console.OutputEncoding = utf8;
Console.SetError(new StreamWriter(Console.OpenStandardError(), utf8) { AutoFlush = true });

if (args.Length > 0 && string.Equals(args[0], "structure", StringComparison.Ordinal))
    return await new StructureCommand().ExecuteAsync(args, Console.Out, Console.Error);

if (args.Length > 0 && string.Equals(args[0], "analyze", StringComparison.OrdinalIgnoreCase))
{
    var command = new AnalyzeCommand(
        new DemoAnalysisService(new WordFormattingReader(inferUnstyledStructure: true), new FontSizeAnalyzer()));
    return command.Execute(args);
}

if (args.Length > 0 && string.Equals(args[0], "analyze-styles", StringComparison.OrdinalIgnoreCase))
{
    var command = new StyleAnalyzeCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "apply-nku-format", StringComparison.OrdinalIgnoreCase))
{
    var command = new ApplyNkuFormatCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "analyze-figures", StringComparison.OrdinalIgnoreCase))
{
    var command = new FigureTableAnalyzeCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "analyze-document", StringComparison.OrdinalIgnoreCase))
{
    var command = new DocumentAnalyzeCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "analyze-formula", StringComparison.OrdinalIgnoreCase))
{
    var command = new FormulaAnalyzeCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "analyze-footnotes", StringComparison.OrdinalIgnoreCase))
{
    var command = new FootnoteAnalyzeCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "analyze-template", StringComparison.OrdinalIgnoreCase))
{
    var command = new AnalyzeTemplateCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "analyze-abstract", StringComparison.OrdinalIgnoreCase))
{
    var command = new AbstractFormatAnalyzeCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "analyze-back-matter", StringComparison.OrdinalIgnoreCase))
{
    var command = new BackMatterAnalyzeCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "report", StringComparison.OrdinalIgnoreCase))
{
    var command = new ComplianceReportCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "analyze-body-text", StringComparison.OrdinalIgnoreCase))
{
    var command = new BodyTextFormatCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "analyze-headings", StringComparison.OrdinalIgnoreCase))
{
    var command = new HeadingFormatCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "apply-figure-format", StringComparison.OrdinalIgnoreCase))
{
    var command = new ApplyFigureFormatCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "apply-formula-format", StringComparison.OrdinalIgnoreCase))
{
    var command = new ApplyFormulaFormatCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "apply-document-format", StringComparison.OrdinalIgnoreCase))
{
    var command = new ApplyDocumentFormatCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "full-format", StringComparison.OrdinalIgnoreCase))
{
    var command = new FullFormatCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "undo", StringComparison.OrdinalIgnoreCase))
{
    var command = new UndoCommand();
    return command.Execute(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "chat", StringComparison.OrdinalIgnoreCase))
{
    var command = new ChatCommand();
    return command.Execute(args[1..]);
}

if (args.Length != 1)
{
    Console.Error.WriteLine("Usage: ThesisEngine.Cli <document.docx>");
    return 2;
}
try
{
    var document = DocumentLoader.Load(args[0]);
    Console.WriteLine("Document loaded successfully");
    Console.WriteLine($"Parts: {document.PartCount}. 占位模式，尚未检测格式；不能据此判断合规。");
    return 0;
}
catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException or ArgumentException or XmlException or NotSupportedException)
{
    Console.Error.WriteLine($"Unable to load document: {ex.Message}");
    return 1;
}
