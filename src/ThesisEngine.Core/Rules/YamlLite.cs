using System.Globalization;
using System.Text.Json;

namespace ThesisEngine.Core.Rules;

/// <summary>
/// A deliberately small YAML reader for the rule files in this repository.
/// It supports mappings, sequences, inline maps/lists, quoted scalars, comments
/// and block scalars. Anchors, aliases and tags are rejected instead of guessed.
/// </summary>
public abstract class YamlNode
{
    public static YamlNode Parse(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        return new YamlParser(text).Parse();
    }
}

public sealed class YamlMap : YamlNode
{
    private readonly Dictionary<string, YamlNode> _values = new(StringComparer.Ordinal);

    public IReadOnlyDictionary<string, YamlNode> Values => _values;

    public void Add(string key, YamlNode value)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(key);
        ArgumentNullException.ThrowIfNull(value);
        if (!_values.TryAdd(key, value))
            throw new FormatException($"YAML duplicate key: {key}");
    }

    public YamlNode? Get(string key) => _values.TryGetValue(key, out var value) ? value : null;

    public YamlMap? GetMap(string key) => Get(key) as YamlMap;

    public YamlSequence? GetSequence(string key) => Get(key) as YamlSequence;

    public string? GetString(string key) => Get(key) is YamlScalar scalar
        ? scalar.Value?.ToString()
        : null;
}

public sealed class YamlSequence : YamlNode
{
    private readonly List<YamlNode> _items = new();

    public IReadOnlyList<YamlNode> Items => _items;

    public void Add(YamlNode value)
    {
        ArgumentNullException.ThrowIfNull(value);
        _items.Add(value);
    }
}

public sealed class YamlScalar : YamlNode
{
    public YamlScalar(object? value, string raw)
    {
        Value = value;
        Raw = raw;
    }

    public object? Value { get; }

    public string Raw { get; }

    public bool IsNull => Value is null;

    public override string ToString() => Value switch
    {
        null => string.Empty,
        bool boolean => boolean ? "true" : "false",
        double number => number.ToString("0.############", CultureInfo.InvariantCulture),
        IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
        _ => Value.ToString() ?? string.Empty
    };
}

internal sealed class YamlParser
{
    private sealed record SourceLine(int Indent, string Content, int Number);

    private readonly List<SourceLine> _lines = new();
    private int _index;

    public YamlParser(string text)
    {
        var normalized = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
        var rawLines = normalized.Split('\n');
        for (var number = 0; number < rawLines.Length; number++)
        {
            var stripped = StripComment(rawLines[number]);
            if (string.IsNullOrWhiteSpace(stripped)) continue;

            var indent = 0;
            while (indent < stripped.Length && stripped[indent] == ' ') indent++;
            if (indent < stripped.Length && stripped[indent] == '\t')
                throw new FormatException($"YAML line {number + 1}: tabs are not valid indentation.");

            var content = stripped[indent..].TrimEnd();
            if (content.Length == 0) continue;
            _lines.Add(new SourceLine(indent, content, number + 1));
        }
    }

    public YamlNode Parse()
    {
        if (_lines.Count == 0) return new YamlMap();
        var node = ParseBlock(_lines[0].Indent);
        if (_index != _lines.Count)
            throw Error($"unexpected trailing content");
        return node;
    }

    private YamlNode ParseBlock(int indent)
    {
        if (_index >= _lines.Count) return new YamlMap();
        var line = _lines[_index];
        if (line.Indent < indent) return new YamlMap();
        if (line.Indent > indent) throw Error("unexpected indentation");
        return IsSequenceLine(line.Content) ? ParseSequence(indent) : ParseMapping(indent);
    }

    private YamlMap ParseMapping(int indent)
    {
        var map = new YamlMap();
        while (_index < _lines.Count)
        {
            var line = _lines[_index];
            if (line.Indent < indent) break;
            if (line.Indent > indent) throw Error("unexpected indentation");
            if (IsSequenceLine(line.Content)) break;
            ParseMapEntry(line, indent, map);
        }
        return map;
    }

    private void ParseMapEntry(SourceLine line, int indent, YamlMap map)
    {
        if (!TrySplitKeyValue(line.Content, out var key, out var valueText))
            throw Error($"expected a mapping entry, got: {line.Content}");

        _index++;
        if (valueText.Length == 0)
        {
            if (_index < _lines.Count && _lines[_index].Indent > indent)
                map.Add(key, ParseBlock(_lines[_index].Indent));
            else
                map.Add(key, new YamlScalar(null, string.Empty));
            return;
        }

        if (valueText is "|" or ">")
        {
            map.Add(key, new YamlScalar(ReadBlockScalar(indent, valueText == ">"), valueText));
            return;
        }

        map.Add(key, ParseInline(valueText));
        if (_index < _lines.Count && _lines[_index].Indent > indent)
            throw Error($"inline value for '{key}' must not have an indented child block");
    }

    private YamlSequence ParseSequence(int indent)
    {
        var sequence = new YamlSequence();
        while (_index < _lines.Count)
        {
            var line = _lines[_index];
            if (line.Indent < indent) break;
            if (line.Indent > indent) throw Error("unexpected indentation in sequence");
            if (!IsSequenceLine(line.Content)) break;

            var rest = line.Content.Length <= 1 ? string.Empty : line.Content[1..].TrimStart();
            if (rest.Length == 0)
            {
                _index++;
                if (_index < _lines.Count && _lines[_index].Indent > indent)
                    sequence.Add(ParseBlock(_lines[_index].Indent));
                else
                    sequence.Add(new YamlScalar(null, string.Empty));
                continue;
            }

            if (IsMappingStart(rest))
            {
                // Re-home "- key: value" as an ordinary mapping line at indent + 2.
                _lines[_index] = new SourceLine(indent + 2, rest, line.Number);
                sequence.Add(ParseMapping(indent + 2));
                continue;
            }

            sequence.Add(ParseInline(rest));
            _index++;
        }
        return sequence;
    }

    private string ReadBlockScalar(int parentIndent, bool folded)
    {
        var lines = new List<string>();
        var blockIndent = -1;
        while (_index < _lines.Count)
        {
            var line = _lines[_index];
            if (line.Indent <= parentIndent) break;
            if (blockIndent < 0) blockIndent = line.Indent;
            lines.Add(line.Indent >= blockIndent ? new string(' ', line.Indent - blockIndent) + line.Content : line.Content);
            _index++;
        }
        return string.Join(folded ? " " : "\n", lines);
    }

    private YamlNode ParseInline(string text)
    {
        var trimmed = text.Trim();
        if (trimmed.Length == 0) return new YamlScalar(null, string.Empty);
        if (trimmed[0] == '{') return ParseInlineMap(trimmed);
        if (trimmed[0] == '[') return ParseInlineSequence(trimmed);
        if (trimmed[0] == '\'') return new YamlScalar(ParseSingleQuoted(trimmed), trimmed);
        if (trimmed[0] == '"') return new YamlScalar(ParseDoubleQuoted(trimmed), trimmed);
        if (trimmed is "null" or "~" or "Null" or "NULL") return new YamlScalar(null, trimmed);
        if (bool.TryParse(trimmed, out var boolean)) return new YamlScalar(boolean, trimmed);
        if (long.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out var integer))
            return new YamlScalar(integer, trimmed);
        if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            return new YamlScalar(number, trimmed);
        if (trimmed.StartsWith('&') || trimmed.StartsWith('*') || trimmed.StartsWith('!'))
            throw Error($"YAML anchors, aliases and tags are not supported: {trimmed}");
        return new YamlScalar(trimmed, trimmed);
    }

    private YamlMap ParseInlineMap(string text)
    {
        if (!text.EndsWith('}')) throw Error($"unterminated inline map: {text}");
        var map = new YamlMap();
        var body = text[1..^1].Trim();
        if (body.Length == 0) return map;
        foreach (var part in SplitTopLevel(body, ','))
        {
            if (!TrySplitKeyValue(part, out var key, out var valueText))
                throw Error($"invalid inline map entry: {part}");
            map.Add(key, valueText.Length == 0 ? new YamlScalar(null, string.Empty) : ParseInline(valueText));
        }
        return map;
    }

    private YamlSequence ParseInlineSequence(string text)
    {
        if (!text.EndsWith(']')) throw Error($"unterminated inline sequence: {text}");
        var sequence = new YamlSequence();
        var body = text[1..^1].Trim();
        if (body.Length == 0) return sequence;
        foreach (var part in SplitTopLevel(body, ','))
            sequence.Add(ParseInline(part));
        return sequence;
    }

    private string ParseSingleQuoted(string text)
    {
        if (text.Length < 2 || text[^1] != '\'') throw Error($"unterminated single-quoted scalar: {text}");
        return text[1..^1].Replace("''", "'", StringComparison.Ordinal);
    }

    private string ParseDoubleQuoted(string text)
    {
        if (text.Length < 2 || text[^1] != '"') throw Error($"unterminated double-quoted scalar: {text}");
        try
        {
            return JsonSerializer.Deserialize<string>(text) ?? string.Empty;
        }
        catch (JsonException)
        {
            return text[1..^1]
                .Replace("\\\"", "\"", StringComparison.Ordinal)
                .Replace("\\n", "\n", StringComparison.Ordinal)
                .Replace("\\t", "\t", StringComparison.Ordinal)
                .Replace("\\\\", "\\", StringComparison.Ordinal);
        }
    }

    private FormatException Error(string message)
    {
        var line = _index < _lines.Count ? _lines[_index] : _lines[^1];
        return new FormatException($"YAML line {line.Number}: {message}");
    }

    private static bool IsSequenceLine(string content) =>
        content == "-" || content.StartsWith("- ", StringComparison.Ordinal);

    private static bool IsMappingStart(string content)
    {
        if (content.Length == 0 || content[0] is '{' or '[' or '\'' or '"') return false;
        return TrySplitKeyValue(content, out _, out _);
    }

    private static bool TrySplitKeyValue(string content, out string key, out string value)
    {
        key = string.Empty;
        value = string.Empty;
        var inSingle = false;
        var inDouble = false;
        var depth = 0;
        for (var index = 0; index < content.Length; index++)
        {
            var character = content[index];
            if (inDouble)
            {
                if (character == '\\') { index++; continue; }
                if (character == '"') inDouble = false;
                continue;
            }
            if (inSingle)
            {
                if (character == '\'')
                {
                    if (index + 1 < content.Length && content[index + 1] == '\'') { index++; continue; }
                    inSingle = false;
                }
                continue;
            }
            switch (character)
            {
                case '\'': inSingle = true; continue;
                case '"': inDouble = true; continue;
                case '{':
                case '[': depth++; continue;
                case '}':
                case ']': depth--; continue;
                case ':' when depth == 0 && (index + 1 == content.Length || char.IsWhiteSpace(content[index + 1])):
                    key = Unquote(content[..index].Trim());
                    value = content[(index + 1)..].Trim();
                    return key.Length > 0;
            }
        }
        return false;
    }

    private static IEnumerable<string> SplitTopLevel(string text, char separator)
    {
        var start = 0;
        var inSingle = false;
        var inDouble = false;
        var depth = 0;
        for (var index = 0; index < text.Length; index++)
        {
            var character = text[index];
            if (inDouble)
            {
                if (character == '\\') { index++; continue; }
                if (character == '"') inDouble = false;
                continue;
            }
            if (inSingle)
            {
                if (character == '\'')
                {
                    if (index + 1 < text.Length && text[index + 1] == '\'') { index++; continue; }
                    inSingle = false;
                }
                continue;
            }
            switch (character)
            {
                case '\'': inSingle = true; break;
                case '"': inDouble = true; break;
                case '{':
                case '[': depth++; break;
                case '}':
                case ']': depth--; break;
                default:
                    if (character == separator && depth == 0)
                    {
                        yield return text[start..index].Trim();
                        start = index + 1;
                    }
                    break;
            }
        }
        yield return text[start..].Trim();
    }

    private static string Unquote(string text)
    {
        if (text.Length >= 2 && text[0] == '\'' && text[^1] == '\'')
            return text[1..^1].Replace("''", "'", StringComparison.Ordinal);
        if (text.Length >= 2 && text[0] == '"' && text[^1] == '"')
        {
            try { return JsonSerializer.Deserialize<string>(text) ?? string.Empty; }
            catch (JsonException) { return text[1..^1]; }
        }
        return text;
    }

    private static string StripComment(string line)
    {
        var inSingle = false;
        var inDouble = false;
        for (var index = 0; index < line.Length; index++)
        {
            var character = line[index];
            if (inDouble)
            {
                if (character == '\\') { index++; continue; }
                if (character == '"') inDouble = false;
                continue;
            }
            if (inSingle)
            {
                if (character == '\'')
                {
                    if (index + 1 < line.Length && line[index + 1] == '\'') { index++; continue; }
                    inSingle = false;
                }
                continue;
            }
            if (character == '\'') { inSingle = true; continue; }
            if (character == '"') { inDouble = true; continue; }
            if (character == '#' && (index == 0 || char.IsWhiteSpace(line[index - 1])))
                return line[..index];
        }
        return line;
    }
}
