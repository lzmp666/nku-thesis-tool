using System;
using System.IO;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO;

// The VSTO designer supplies the AddInBase partial type and lifecycle events
// when this project is opened by the Office/SharePoint Visual Studio workload.
public partial class ThisAddIn
{
    private EngineProtocolClient? _engineClient;

    public bool IsEngineConfigured => _engineClient is not null;

    public EngineResponse AnalyzeDocument(string inputPath)
    {
        if (_engineClient is null)
        {
            return new EngineResponse
            {
                ProtocolVersion = EngineProtocol.Version,
                Ok = false,
                Error = new EngineError
                {
                    Code = EngineProtocol.EngineExit,
                    Message = "引擎未配置。请设置 THESIS_ENGINE_PROTOCOL_HOST 为协议宿主 DLL 路径。"
                }
            };
        }

        return _engineClient.Analyze(inputPath);
    }

    private void ThisAddIn_Startup(object sender, EventArgs e)
    {
        var hostPath = Environment.GetEnvironmentVariable("THESIS_ENGINE_PROTOCOL_HOST");
        if (!string.IsNullOrWhiteSpace(hostPath) && File.Exists(hostPath))
            _engineClient = new EngineProtocolClient(EngineProcessOptions.ForDotnet(hostPath));
    }

    private void ThisAddIn_Shutdown(object sender, EventArgs e)
    {
        _engineClient = null;
    }

    #region VSTO generated code hook

    private void InternalStartup()
    {
        Startup += ThisAddIn_Startup;
        Shutdown += ThisAddIn_Shutdown;
    }

    #endregion
}
