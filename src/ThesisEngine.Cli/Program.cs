using System.Xml;
using ThesisEngine.Core;

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
