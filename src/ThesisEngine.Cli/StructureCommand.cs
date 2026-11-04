using System.Text;
using System.Text.Json;
using ThesisEngine.Core;

namespace ThesisEngine.Cli;

/// <summary>An explicit, local structure review/formatting workflow; analysis alone never writes a DOCX.</summary>
public sealed class StructureCommand
{
    public async Task<int> ExecuteAsync(string[] arguments, TextWriter output, TextWriter error,
        CancellationToken cancellationToken = default)
    {
        if (!TryParse(arguments, out var options))
        {
            error.WriteLine("用法：ThesisEngine.Cli structure <file.docx> [--report new.json] [--approved-outline outline.json] " +
                "[--ollama-model name --local-model-confirmed] [--ollama-endpoint http://127.0.0.1:11434/api/generate] " +
                "[--format-output new.docx --nku2026-draft mode1|mode2] [--training-output new.jsonl]");
            return 2;
        }
        try
        {
            var input = Path.GetFullPath(options.Input);
            if (!File.Exists(input) || !string.Equals(Path.GetExtension(input), ".docx", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("结构工作流输入必须是存在的 .docx 文件。");
            var outputs = new[] { options.Report, options.FormattedOutput, options.TrainingOutput }
                .Where(path => path is not null).Select(path => Path.GetFullPath(path!)).ToArray();
            var comparer = OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
            if (outputs.Distinct(comparer).Count() != outputs.Length || outputs.Contains(input, comparer))
                throw new ArgumentException("输出必须彼此不同且不能是输入原文件。");
            foreach (var path in outputs)
            {
                if (File.Exists(path) || Directory.Exists(path)) throw new IOException("输出已存在，未覆盖：" + path);
                if (!Directory.Exists(Path.GetDirectoryName(path))) throw new DirectoryNotFoundException("请先创建输出目录。");
            }
            if (options.FormattedOutput is not null && !string.Equals(Path.GetExtension(options.FormattedOutput), ".docx", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("排版输出必须是 .docx。");
            var snapshot = new WordFormattingReader(inferUnstyledStructure: true,
                includePackageFingerprint: options.FormattedOutput is not null).Read(input);
            ThesisOutlineApproval? approval = null;
            if (options.ApprovedOutline is not null)
            {
                approval = ThesisStructureJson.DeserializeStrict<ThesisOutlineApproval>(ReadBounded(options.ApprovedOutline));
                snapshot = ThesisOutlineReview.Apply(snapshot, approval);
            }
            ThesisAiStructureReport? ai = null;
            if (options.Model is not null)
            {
                using var model = new OllamaThesisStructureModel(options.Model,
                    options.Endpoint is null ? null : new Uri(options.Endpoint), localOnlyServiceConfirmed: options.LocalModelConfirmed);
                ai = await new ThesisStructureAiCoordinator(model).ProposeAsync(snapshot, cancellationToken);
            }

            ThesisRoleFormatPlan? formattingPlan = null;
            ThesisRoleFormatApplyResult? formatted = null;
            if (options.FormattedOutput is not null)
            {
                var mode = options.TitleMode == "mode1" ? ThesisTitleMode.Mode1 : ThesisTitleMode.Mode2;
                var formatter = new ThesisRoleFormatter();
                formattingPlan = formatter.Plan(snapshot, ThesisRoleFormattingProfile.Nku2026Draft(mode), authorSelectedDraft: true);
                formatted = formatter.Apply(input, formattingPlan, options.FormattedOutput);
            }
            if (options.TrainingOutput is not null)
            {
                if (approval is null) throw new ArgumentException("训练数据只能从 --approved-outline 的作者确认标签导出。");
                var samples = ThesisOutlineReview.ExportApprovedSamples(snapshot, approval);
                WriteNew(options.TrainingOutput, string.Join(Environment.NewLine,
                    samples.Select(sample => JsonSerializer.Serialize(sample, ThesisStructureJson.Options))) + Environment.NewLine);
            }
            var fingerprint = ThesisStructureAiCoordinator.Fingerprint(snapshot);
            var report = new
            {
                schemaVersion = "thesis-structure-preview-1", documentFingerprint = fingerprint,
                source = input, inputChanged = false,
                summary = new
                {
                    blocks = snapshot.Structure!.Blocks.Count,
                    review = snapshot.Structure.Blocks.Count(block => block.RequiresReview || block.Confidence < 0.90),
                    recognizedBody = snapshot.Paragraphs.Count(paragraph => paragraph.Scope == ParagraphScopeClassification.ConfirmedBodyProse)
                },
                structure = snapshot.Structure,
                sourceBlocks = snapshot.SourceMap!.Blocks,
                approvalTemplate = new ThesisOutlineApproval(fingerprint, snapshot.Structure.Blocks
                    .Where(block => (block.RequiresReview || block.Confidence < 0.90) &&
                        snapshot.SourceMap.Blocks[block.ParagraphIndex].HasCompleteTextMapping &&
                        !snapshot.Paragraphs[block.ParagraphIndex].HasUnsupportedInlineContent &&
                        !snapshot.Paragraphs[block.ParagraphIndex].IsTableCell)
                    .Select(block => new ApprovedThesisBlock(block.BlockId, block.Section, block.Kind, block.HeadingLevel)).ToArray()),
                ai, formattingPlan, formatted,
                limitations = new[]
                {
                    "结构分数是证据强度，非已校准的正确率；AI 结果须确认大纲。",
                    "南开 2026 角色排版是作者显式选择的 draft 预览，不是整篇合规结论。",
                    "未覆盖的封面、自动目录、页眉页码、表格单元格、复杂图表和公式保留原样并报告。"
                }
            };
            var json = JsonSerializer.Serialize(report, new JsonSerializerOptions(ThesisStructureJson.Options) { WriteIndented = true });
            if (options.Report is null) await output.WriteLineAsync(json);
            else { WriteNew(options.Report, json); await output.WriteLineAsync("结构报告：" + Path.GetFullPath(options.Report)); }
            return 0;
        }
        catch (Exception exception) when (exception is IOException or ArgumentException or JsonException or InvalidOperationException or FormatException or
            NotSupportedException or UnauthorizedAccessException or System.Xml.XmlException)
        {
            await error.WriteLineAsync("结构工作流未完成：" + exception.Message);
            return 1;
        }
    }

    private static string ReadBounded(string path)
    {
        var info = new FileInfo(path);
        if (info.Length > 1_048_576) throw new InvalidDataException("确认大纲文件超过大小限制。");
        return File.ReadAllText(path, Encoding.UTF8);
    }

    private static void WriteNew(string path, string text)
    {
        using var stream = new FileStream(Path.GetFullPath(path), FileMode.CreateNew, FileAccess.Write, FileShare.None);
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(text);
    }

    private sealed record Options(string Input, string? Report, string? ApprovedOutline, string? Model,
        string? Endpoint, bool LocalModelConfirmed, string? FormattedOutput, string? TitleMode, string? TrainingOutput);

    private static bool TryParse(string[] args, out Options options)
    {
        options = new Options(string.Empty, null, null, null, null, false, null, null, null);
        if (args.Length < 2 || args[0] != "structure" || string.IsNullOrWhiteSpace(args[1])) return false;
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        var local = false;
        for (var index = 2; index < args.Length; index++)
        {
            var name = args[index];
            if (name == "--local-model-confirmed") { if (local) return false; local = true; continue; }
            if (name is not ("--report" or "--approved-outline" or "--ollama-model" or "--ollama-endpoint" or
                "--format-output" or "--nku2026-draft" or "--training-output") ||
                ++index >= args.Length || string.IsNullOrWhiteSpace(args[index]) || args[index].StartsWith("--", StringComparison.Ordinal) ||
                !values.TryAdd(name, args[index])) return false;
        }
        string? Get(string key) => values.GetValueOrDefault(key);
        var format = Get("--format-output"); var mode = Get("--nku2026-draft");
        if ((format is null) != (mode is null) || mode is not (null or "mode1" or "mode2") ||
            Get("--training-output") is not null && Get("--approved-outline") is null ||
            Get("--ollama-model") is null && (local || Get("--ollama-endpoint") is not null)) return false;
        options = new Options(args[1], Get("--report"), Get("--approved-outline"), Get("--ollama-model"),
            Get("--ollama-endpoint"), local, format, mode, Get("--training-output"));
        return true;
    }
}
