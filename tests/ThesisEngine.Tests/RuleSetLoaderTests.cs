using ThesisEngine.Core.Rules;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class RuleSetLoaderTests
{
    [Fact]
    public void GeneralProfileLoadsEveryCatalogSlotWithoutWarnings()
    {
        var ruleSet = RuleSetLoader.Load("nku-graduate");

        Assert.Equal("nku-graduate", ruleSet.ProfileId);
        Assert.Equal("general-draft", ruleSet.RuleVersion);
        Assert.Equal("draft", ruleSet.Status);
        Assert.True(ruleSet.IsDraft);
        Assert.Empty(ruleSet.Warnings);
        Assert.True(ruleSet.Rules.Count >= 100);
        Assert.Equal(ruleSet.Rules.Count, ruleSet.Rules.Select(rule => rule.RuleId).Distinct(StringComparer.Ordinal).Count());
        Assert.All(ruleSet.Rules, rule =>
        {
            Assert.False(string.IsNullOrWhiteSpace(rule.RuleId));
            Assert.StartsWith("rules.", rule.YamlPath, StringComparison.Ordinal);
            Assert.False(string.IsNullOrWhiteSpace(rule.Clause));
        });
    }

    [Fact]
    public void GeneralProfileMarksImplementedAndUnimplementedRulesExplicitly()
    {
        var ruleSet = RuleSetLoader.Load("nku-graduate");

        var bodyFont = ruleSet.Rules.Single(rule => rule.RuleId == "R11/font.cn");
        Assert.True(bodyFont.Implemented);
        Assert.Equal("宋体", bodyFont.ExpectedValue);
        Assert.Equal("4.7 表4.6", bodyFont.Clause);
        Assert.Equal("rules.body_text.font.cn", bodyFont.YamlPath);

        var headerFont = ruleSet.Rules.Single(rule => rule.RuleId == "R2/font_cn");
        Assert.False(headerFont.Implemented);
        Assert.False(string.IsNullOrWhiteSpace(headerFont.UncheckedReason));

        var margin = ruleSet.Rules.Single(rule => rule.RuleId == "R1/margins_cm");
        Assert.True(margin.Implemented);
        Assert.Contains("3.8", margin.ExpectedDisplay, StringComparison.Ordinal);
        Assert.Contains("3.2", margin.ExpectedDisplay, StringComparison.Ordinal);
    }

    [Fact]
    public void EveryUnimplementedRuleHasAReason()
    {
        var ruleSet = RuleSetLoader.Load("nku-graduate");
        Assert.All(ruleSet.Rules.Where(rule => !rule.Implemented),
            rule => Assert.False(string.IsNullOrWhiteSpace(rule.UncheckedReason)));
    }

    [Fact]
    public void LegacyReportSwitchCanSuppressOneGroupWithoutHidingIt()
    {
        var original = File.ReadAllText(RuleSetLoader.FindDefaultRulesFile("nku-graduate"));
        var path = WriteTempRuleSet(original.Replace("suppress: []", "suppress: [body_text]", StringComparison.Ordinal));
        try
        {
            var ruleSet = RuleSetLoader.Load("nku-graduate", path);
            Assert.False(ruleSet.Rules.Single(rule => rule.RuleId == "R11/font.cn").Report);
            Assert.True(ruleSet.Rules.Single(rule => rule.RuleId == "R1/paper_size").Report);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Flat2026ProfileLoadsThreeDraftBodyRules()
    {
        var ruleSet = RuleSetLoader.Load("nku-graduate-2026-body");

        Assert.Equal("nku-graduate-2026-body", ruleSet.ProfileId);
        Assert.Equal("nku-graduate-2026-body-v1", ruleSet.RuleVersion);
        Assert.Equal("draft", ruleSet.Status);
        Assert.Equal(3, ruleSet.Rules.Count);
        Assert.All(ruleSet.Rules, rule =>
        {
            Assert.StartsWith("NKU-REQ-", rule.RuleId, StringComparison.Ordinal);
            Assert.StartsWith("M26-", rule.MatrixId, StringComparison.Ordinal);
            Assert.True(rule.Implemented);
        });
        Assert.Contains(ruleSet.Rules, rule => rule.RuleId == "NKU-REQ-042" && rule.ExpectedValue == "宋体");
        Assert.Contains(ruleSet.Rules, rule => rule.RuleId == "NKU-REQ-043" && rule.ExpectedValue == "Times New Roman");
        Assert.Contains(ruleSet.Rules, rule => rule.RuleId == "NKU-REQ-044" && rule.ExpectedValue == "12");
    }

    [Fact]
    public void CustomFlatRuleFileIsLoaded()
    {
        var path = WriteTempRuleSet("""
            schema_version: "0.2.0"
            status: draft
            profile:
              profile_id: custom-test
              rule_version: custom-1
              standard_version: "2026"
              status: draft
              selection: explicit
              scope_policy:
                include: [confirmed_body_prose]
                exclude: []
                unknown: unsupported
              source:
                title: test
                url: "https://example.com/spec"
                version: "2026"
                locator: "4.7"
                quote: "test"
                verified_on: "2026-01-01"
            rules:
              - id: NKU-REQ-901
                matrix_id: M26-901
                scope: [body]
                property: font_size
                unknown_policy: unsupported
                expect: {value: 12, unit: pt}
                severity: warning
                fix: manual
                source:
                  title: test
                  url: "https://example.com/spec"
                  version: "2026"
                  locator: "第4.7节表4.6"
                  quote: "test"
                  verified_on: "2026-01-01"
            """);
        try
        {
            var ruleSet = RuleSetLoader.Load("custom-test", path);
            Assert.Single(ruleSet.Rules);
            var rule = ruleSet.Rules[0];
            Assert.Equal("NKU-REQ-901", rule.RuleId);
            Assert.Equal("M26-901", rule.MatrixId);
            Assert.Equal("第4.7节表4.6", rule.Clause);
            Assert.Equal("12pt", rule.ExpectedDisplay);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void YamlParserHandlesInlineMapsListsAndComments()
    {
        var root = YamlNode.Parse("""
            # comment
            meta: {profile: nku-graduate, status: draft} # inline
            list: [a, "b c", 3]
            rules:
              body_text:
                id: R11
                font: {cn: 宋体, western: Times New Roman}
                spacing: {before_pt: 0, after_pt: 0}
            """) as YamlMap;

        Assert.NotNull(root);
        var meta = root.GetMap("meta");
        Assert.NotNull(meta);
        Assert.Equal("nku-graduate", meta.GetString("profile"));
        Assert.Equal(3, root.GetSequence("list")!.Items.Count);
        var rules = root.GetMap("rules")!;
        var body = rules.GetMap("body_text")!;
        Assert.Equal("R11", body.GetString("id"));
        Assert.Equal("宋体", body.GetMap("font")!.GetString("cn"));
        Assert.Equal("0", body.GetMap("spacing")!.GetString("before_pt"));
    }

    [Fact]
    public void MalformedYamlFailsClosed()
    {
        Assert.Throws<FormatException>(() => YamlNode.Parse("""
            rules:
              body_text:
                id: R11
               font: {cn: 宋体}
            """));
    }

    private static string WriteTempRuleSet(string content)
    {
        var path = Path.Combine(Path.GetTempPath(), $"thesis-rules-{Guid.NewGuid():N}.yaml");
        File.WriteAllText(path, content);
        return path;
    }
}
