using System.Globalization;
using System.Text;

namespace ThesisEngine.Core.Rules;

/// <summary>
/// Loads either the legacy `general.yaml` shape or the flat 0.2.0 rule array.
/// Legacy groups are mapped through <see cref="LegacyRuleCatalog"/>; anything the
/// catalog does not implement is emitted as 未检查 instead of disappearing.
/// </summary>
public static class RuleSetLoader
{
    public const string DefaultProfile = "nku-graduate";

    public static AnalysisRuleSet Load(string profile = DefaultProfile, string? ruleFile = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(profile);
        var path = string.IsNullOrWhiteSpace(ruleFile)
            ? FindDefaultRulesFile(profile)
            : Path.GetFullPath(ruleFile);
        if (!File.Exists(path))
            throw new FileNotFoundException($"找不到规则文件：{path}", path);

        var root = YamlNode.Parse(File.ReadAllText(path, Encoding.UTF8)) as YamlMap
            ?? throw new InvalidDataException($"规则文件根节点必须是映射：{path}");
        return root.Get("schema_version") is not null
            ? LoadFlat(root, path, profile)
            : LoadLegacy(root, path, profile);
    }

    public static string FindDefaultRulesFile(string profile)
    {
        var relative = profile switch
        {
            "nku-graduate" => Path.Combine("rules", "nku-graduate", "general.yaml"),
            "nku-graduate-2026-body" => Path.Combine("rules", "nku-graduate", "2026-body-v1.yaml"),
            _ => Path.Combine("rules", profile, "general.yaml")
        };

        foreach (var start in new[] { Directory.GetCurrentDirectory(), AppContext.BaseDirectory })
        {
            var directory = new DirectoryInfo(start);
            for (var depth = 0; directory is not null && depth < 12; depth++, directory = directory.Parent)
            {
                var candidate = Path.Combine(directory.FullName, relative);
                if (File.Exists(candidate)) return candidate;
            }
        }

        throw new FileNotFoundException(
            $"找不到 profile「{profile}」的规则文件；请从仓库根目录运行，或使用 --rule-file 指定文件。");
    }

    private static AnalysisRuleSet LoadLegacy(YamlMap root, string path, string requestedProfile)
    {
        var meta = root.GetMap("meta");
        var rulesNode = root.GetMap("rules")
            ?? throw new InvalidDataException("规则文件缺少 rules 映射。");
        var status = meta?.GetString("status") ?? "draft";
        var profileId = meta?.GetString("profile") ?? requestedProfile;
        var fileId = meta?.GetString("file") ?? Path.GetFileNameWithoutExtension(path);
        var ruleVersion = string.Equals(status, "approved", StringComparison.OrdinalIgnoreCase)
            ? fileId
            : fileId + "-draft";

        var rules = new List<AnalysisRule>();
        var warnings = new List<string>();
        foreach (var slot in LegacyRuleCatalog.Slots)
        {
            var groupNode = ResolveNode(rulesNode, slot.GroupPath);
            if (groupNode is null)
            {
                warnings.Add($"规则文件缺少规则组 rules.{slot.GroupPath}；对应条目保留为未检查。");
                rules.Add(new AnalysisRule(
                    slot.GroupPath + (slot.RelativePath.Length == 0 ? string.Empty : "/" + slot.RelativePath),
                    slot.GroupPath,
                    "rules." + slot.GroupPath + (slot.RelativePath.Length == 0 ? string.Empty : "." + slot.RelativePath),
                    "未提供",
                    "未提供",
                    slot.ObjectKind,
                    slot.Property,
                    "(规则文件未给出)",
                    string.Empty,
                    "warning",
                    "manual",
                    false,
                    slot.UncheckedReason ?? "规则文件缺少该规则组",
                    0.1));
                continue;
            }

            var expectedNode = slot.RelativePath.Length == 0
                ? groupNode
                : ResolveNode(groupNode, slot.RelativePath);
            var groupId = ResolveGroupId(rulesNode, slot.GroupPath, groupNode);
            var (sourceId, clause) = FindAncestorSource(rulesNode, slot.GroupPath) ?? ("未提供", "未提供");
            var implemented = slot.Implemented && expectedNode is not null;
            var reason = implemented
                ? null
                : slot.UncheckedReason ?? (expectedNode is null ? "规则文件缺少期望值" : null);
            var expectedDisplay = slot.ExpectedOverride ?? FormatExpected(slot.Property, expectedNode);
            var expectedValue = NormalizeExpected(slot.Property, expectedNode);
            var baseConfidence = string.Equals(sourceId, "S2026", StringComparison.OrdinalIgnoreCase) ? 0.9 : 0.65;
            var report = IsReportEnabled(root, groupNode, slot.GroupPath);

            rules.Add(new AnalysisRule(
                groupId + (slot.RelativePath.Length == 0 ? string.Empty : "/" + slot.RelativePath),
                groupId,
                "rules." + slot.GroupPath + (slot.RelativePath.Length == 0 ? string.Empty : "." + slot.RelativePath),
                clause,
                sourceId,
                slot.ObjectKind,
                slot.Property,
                expectedDisplay,
                expectedValue,
                slot.Severity,
                slot.FixLevel,
                implemented,
                reason,
                baseConfidence)
            {
                Report = report
            });
        }

        return new AnalysisRuleSet(
            profileId,
            ruleVersion,
            status,
            path,
            rules.AsReadOnly(),
            warnings.AsReadOnly());
    }

    private static AnalysisRuleSet LoadFlat(YamlMap root, string path, string requestedProfile)
    {
        var profile = root.GetMap("profile");
        var status = root.GetString("status") ?? profile?.GetString("status") ?? "draft";
        var profileId = profile?.GetString("profile_id") ?? requestedProfile;
        var ruleVersion = profile?.GetString("rule_version") ?? Path.GetFileNameWithoutExtension(path);
        var array = root.GetSequence("rules")
            ?? throw new InvalidDataException("0.2.0 规则文件缺少 rules 数组。");

        var rules = new List<AnalysisRule>();
        var warnings = new List<string>();
        for (var index = 0; index < array.Items.Count; index++)
        {
            if (array.Items[index] is not YamlMap item)
            {
                warnings.Add($"rules[{index}] 不是映射，已跳过。");
                continue;
            }

            var id = item.GetString("id") ?? $"UNKNOWN-{index + 1}";
            var propertyName = item.GetString("property") ?? string.Empty;
            var scope = item.GetSequence("scope")?.Items
                .OfType<YamlScalar>()
                .Select(scalar => scalar.Value?.ToString() ?? string.Empty)
                .ToArray() ?? Array.Empty<string>();
            var objectKind = ResolveFlatObject(propertyName, scope);
            var (property, scriptSupported) = ResolveFlatProperty(item, propertyName);
            var expect = item.GetMap("expect");
            var valueNode = expect?.Get("value");
            var unit = expect?.GetString("unit") ?? string.Empty;
            var expectedDisplay = FormatFlatExpected(property, valueNode, unit);
            var expectedValue = NormalizeFlatExpected(property, valueNode, unit);
            var source = item.GetMap("source");
            var clause = source?.GetString("locator") ?? "未提供";
            var severity = item.GetString("severity") ?? "warning";
            var fix = item.GetString("fix") ?? "manual";
            var implemented = scriptSupported && objectKind is not null && valueNode is not null;
            var reason = implemented
                ? null
                : !scriptSupported
                    ? "字体脚本槽位或属性尚未实现"
                    : objectKind is null
                        ? "规则 scope 尚未实现"
                        : "规则文件缺少 expect.value";

            rules.Add(new AnalysisRule(
                id,
                id,
                $"rules[{index}]",
                clause,
                source?.GetString("title") ?? "0.2.0",
                objectKind ?? "unsupported",
                property,
                expectedDisplay,
                expectedValue,
                severity,
                fix,
                implemented,
                reason,
                string.Equals(status, "approved", StringComparison.OrdinalIgnoreCase) ? 0.95 : 0.8,
                item.GetString("matrix_id"))
            {
                Report = item.Get("report") is not YamlScalar reportScalar || reportScalar.Value is not false
            });
        }

        return new AnalysisRuleSet(
            profileId,
            ruleVersion,
            status,
            path,
            rules.AsReadOnly(),
            warnings.AsReadOnly());
    }

    private static string? ResolveFlatObject(string property, IReadOnlyList<string> scope)
    {
        if (string.Equals(property, "page_margins", StringComparison.Ordinal)) return AnalysisObjectKinds.Page;
        if (scope.Contains("body", StringComparer.Ordinal)) return AnalysisObjectKinds.BodyText;
        if (scope.Contains("heading", StringComparer.Ordinal)) return AnalysisObjectKinds.HeadingChapter;
        if (scope.Contains("section", StringComparer.Ordinal)) return AnalysisObjectKinds.Page;
        return null;
    }

    private static (string Property, bool Supported) ResolveFlatProperty(YamlMap item, string propertyName)
    {
        var script = item.GetString("text_script");
        return propertyName switch
        {
            "font" when string.Equals(script, "east_asian", StringComparison.OrdinalIgnoreCase) =>
                (RuleProperties.FontEastAsia, true),
            "font" when string.Equals(script, "latin", StringComparison.OrdinalIgnoreCase) =>
                (RuleProperties.FontLatin, true),
            "font_size" => (RuleProperties.FontSize, true),
            "line_spacing" => (RuleProperties.LineSpacing, true),
            "page_margins" => (RuleProperties.PageMargins, true),
            _ => (RuleProperties.Unknown, false)
        };
    }

    private static string FormatFlatExpected(string property, YamlNode? valueNode, string unit)
    {
        if (valueNode is null) return "(规则文件未给出)";
        var value = valueNode is YamlScalar scalar ? scalar.ToString() : valueNode is YamlMap map ? JoinMap(map) : string.Empty;
        return property switch
        {
            RuleProperties.FontSize => value + "pt",
            RuleProperties.LineSpacing when unit == "multiple" => "多倍行距 " + value,
            RuleProperties.LineSpacing when unit == "pt" => "固定值 " + value + "pt",
            RuleProperties.PageMargins => "页边距 " + value + "cm",
            _ => value
        };
    }

    private static string NormalizeFlatExpected(string property, YamlNode? valueNode, string unit)
    {
        if (valueNode is null) return string.Empty;
        if (valueNode is YamlScalar scalar)
        {
            return property switch
            {
                RuleProperties.LineSpacing when unit == "multiple" => "multiple:" + scalar,
                RuleProperties.LineSpacing when unit == "pt" => "exact:" + scalar,
                _ => scalar.ToString()
            };
        }
        if (valueNode is YamlMap map && property == RuleProperties.PageMargins)
            return string.Join(",", new[] { "top", "bottom", "left", "right" }
                .Select(key => map.GetString(key) ?? "0"));
        return valueNode is YamlMap map2 ? JoinMap(map2) : string.Empty;
    }

    private static string FormatExpected(string property, YamlNode? node)
    {
        if (node is null) return "(规则文件未给出)";
        return property switch
        {
            RuleProperties.PageMargins when node is YamlMap margins =>
                $"上{Display(margins, "top")}cm / 下{Display(margins, "bottom")}cm / 左{Display(margins, "left")}cm / 右{Display(margins, "right")}cm",
            RuleProperties.PageHeaderFooter when node is YamlMap headerFooter =>
                $"页眉{Display(headerFooter, "header")}cm / 页脚{Display(headerFooter, "footer")}cm",
            RuleProperties.LineSpacing when node is YamlMap spacing =>
                FormatLineSpacing(spacing),
            RuleProperties.Alignment when node is YamlMap alignment =>
                $"模式一={Display(alignment, "mode1")} / 模式二={Display(alignment, "mode2")}",
            RuleProperties.FontSize when node is YamlScalar scalar => scalar + "pt",
            RuleProperties.SpaceBefore when node is YamlScalar scalar => scalar + "pt",
            RuleProperties.SpaceAfter when node is YamlScalar scalar => scalar + "pt",
            RuleProperties.FirstLineIndent when node is YamlScalar scalar => scalar + " 字符",
            RuleProperties.Bold when node is YamlScalar scalar => scalar.Value is true ? "加粗" : "不加粗",
            RuleProperties.Alignment when node is YamlScalar scalar => DisplayAlignment(scalar.Value?.ToString()),
            _ => node switch
            {
                YamlScalar scalar => scalar.ToString(),
                YamlMap map => JoinMap(map),
                YamlSequence sequence => $"{sequence.Items.Count} 项",
                _ => string.Empty
            }
        };
    }

    private static string NormalizeExpected(string property, YamlNode? node)
    {
        if (node is null) return string.Empty;
        if (node is YamlScalar scalar)
        {
            return property switch
            {
                RuleProperties.Bold => scalar.Value is true ? "true" : "false",
                RuleProperties.Alignment => NormalizeAlignment(scalar.Value?.ToString()),
                _ => scalar.ToString()
            };
        }
        if (node is YamlMap map)
        {
            if (property == RuleProperties.LineSpacing)
            {
                var rule = map.GetString("rule") ?? "single";
                return string.Equals(rule, "exact", StringComparison.OrdinalIgnoreCase)
                    ? "exact:" + (map.GetString("value_pt") ?? "0")
                    : rule;
            }
            if (property == RuleProperties.PageMargins)
                return string.Join(",", new[] { "top", "bottom", "left", "right" }
                    .Select(key => map.GetString(key) ?? "0"));
            return JoinMap(map);
        }
        return node is YamlSequence sequence ? sequence.Items.Count.ToString(CultureInfo.InvariantCulture) : string.Empty;
    }

    private static string FormatLineSpacing(YamlMap map)
    {
        var rule = map.GetString("rule") ?? "single";
        return string.Equals(rule, "exact", StringComparison.OrdinalIgnoreCase)
            ? $"固定值 {map.GetString("value_pt") ?? "?"}pt"
            : string.Equals(rule, "single", StringComparison.OrdinalIgnoreCase)
                ? "单倍行距"
                : rule;
    }

    private static string Display(YamlMap map, string key) => map.GetString(key) ?? "?";

    private static string JoinMap(YamlMap map) =>
        string.Join("; ", map.Values.Select(pair => pair.Key + "=" + pair.Value switch
        {
            YamlScalar scalar => scalar.ToString(),
            YamlMap nested => JoinMap(nested),
            YamlSequence sequence => sequence.Items.Count.ToString(CultureInfo.InvariantCulture),
            _ => string.Empty
        }));

    private static string DisplayAlignment(string? value) => NormalizeAlignment(value) switch
    {
        "left" => "左对齐",
        "center" => "居中",
        "right" => "右对齐",
        "justify" => "两端对齐",
        _ => value ?? "?"
    };

    public static string NormalizeAlignment(string? value) => value?.Trim().ToLowerInvariant() switch
    {
        "both" or "distribute" or "justify" or "justified" => "justify",
        "start" or "left" => "left",
        "end" or "right" => "right",
        "center" => "center",
        _ => string.Empty
    };

    private static string ResolveGroupId(YamlMap rulesNode, string groupPath, YamlNode groupNode)
    {
        if (groupNode is YamlMap groupMap && groupMap.GetString("id") is { Length: > 0 } direct)
            return direct;

        var segments = groupPath.Split('.');
        for (var length = segments.Length - 1; length > 0; length--)
        {
            var parent = ResolveNode(rulesNode, string.Join('.', segments.Take(length)));
            if (parent is YamlMap parentMap && parentMap.GetString("id") is { Length: > 0 } parentId)
                return parentId + "." + string.Join('.', segments.Skip(length));
        }
        return groupPath;
    }

    private static bool IsReportEnabled(YamlMap root, YamlNode groupNode, string groupPath)
    {
        if (groupNode is YamlMap groupMap && groupMap.Get("report") is YamlScalar groupReport &&
            groupReport.Value is bool groupEnabled)
            return groupEnabled;

        var report = root.GetMap("meta")?.GetMap("analysis")?.GetMap("report");
        var defaultEnabled = report?.Get("default") is YamlScalar defaultScalar &&
                             defaultScalar.Value is bool defaultValue
            ? defaultValue
            : true;
        var suppress = report?.GetSequence("suppress")?.Items
            .OfType<YamlScalar>()
            .Select(scalar => scalar.Value?.ToString() ?? string.Empty)
            .Where(value => value.Length > 0)
            .ToArray() ?? Array.Empty<string>();
        return defaultEnabled && !suppress.Any(prefix =>
            string.Equals(groupPath, prefix, StringComparison.Ordinal) ||
            groupPath.StartsWith(prefix + ".", StringComparison.Ordinal));
    }

    private static (string SourceId, string Clause)? FindAncestorSource(YamlMap rulesNode, string groupPath)
    {
        var segments = groupPath.Split('.');
        for (var length = segments.Length; length > 0; length--)
        {
            var current = ResolveNode(rulesNode, string.Join('.', segments.Take(length)));
            if (current is not YamlMap map) continue;
            var source = map.GetMap("source");
            if (source is null) continue;
            return (source.GetString("src") ?? "未提供", source.GetString("clause") ?? "未提供");
        }
        return null;
    }

    private static YamlNode? ResolveNode(YamlNode root, string dottedPath)
    {
        YamlNode? current = root;
        foreach (var segment in dottedPath.Split('.', StringSplitOptions.RemoveEmptyEntries))
        {
            if (current is not YamlMap map) return null;
            current = map.Get(segment);
            if (current is null) return null;
        }
        return current;
    }
}
