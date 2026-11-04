using ThesisAddIn.VSTO.Engine;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class EngineHostPathResolverTests
{
    [Fact]
    public void CurrentUserInstallWinsOverStaleLauncherEnvironment()
    {
        Assert.Equal(@"D:\WordDemo\engine\ThesisEngine.ProtocolHost.dll",
            EngineHostPathResolver.Select(@"D:\WordDemo\engine\ThesisEngine.ProtocolHost.dll", @"C:\old-build\host.dll"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" \t ")]
    public void ProcessConfigurationWorksWhenThereIsNoCurrentUserInstall(string? userHost)
    {
        Assert.Equal(@"D:\developer\host.dll", EngineHostPathResolver.Select(userHost!, @"D:\developer\host.dll"));
    }

    [Fact]
    public void ExplicitUserPathIsNotSilentlyReplacedWithAnOlderHost()
    {
        Assert.Equal(@"D:\missing-install\host.dll",
            EngineHostPathResolver.Select(@"D:\missing-install\host.dll", @"C:\old-build\host.dll"));
    }
}
