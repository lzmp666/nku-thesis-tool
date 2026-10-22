using System.Xml;
using System.Text;
using ThesisEngine.Cli;
using ThesisEngine.Core;

var utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
Console.OutputEncoding = utf8;
Console.SetError(new StreamWriter(Console.OpenStandardError(), utf8) { AutoFlush = true });

if (args.Length > 0 && string.Equals(args[0], "analyze", StringComparison.OrdinalIgnoreCase))
{
    var command = new AnalyzeCommand(
        new DemoAnalysisService(new WordFormattingReader(), new FontSizeAnalyzer()));
    return command.Execute(args);
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
