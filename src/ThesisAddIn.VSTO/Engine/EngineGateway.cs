using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ThesisAddIn.Protocol;

namespace ThesisAddIn.VSTO.Engine
{
    /// <summary>
    /// 插件与引擎进程之间的唯一入口：负责定位契约宿主、发起调用，
    /// 并把协议错误码翻译成面向使用者的中文说明。
    /// </summary>
    public sealed class EngineGateway
    {
        public const string HostPathVariable = "THESIS_ENGINE_PROTOCOL_HOST";

        private readonly EngineProtocolClient _client;

        private EngineGateway(EngineProtocolClient client)
        {
            _client = client;
        }

        public static EngineGateway Create(out string error)
        {
            error = null;
            string userHostPath;
            try
            {
                userHostPath = Environment.GetEnvironmentVariable(HostPathVariable, EnvironmentVariableTarget.User);
            }
            catch (Exception exception)
            {
                error = "无法读取当前用户的引擎配置，已停止连接：" + exception.Message;
                return null;
            }
            string hostPath = EngineHostPathResolver.Select(userHostPath, Environment.GetEnvironmentVariable(HostPathVariable));
            if (string.IsNullOrWhiteSpace(hostPath))
            {
                error = "尚未配置引擎。请把环境变量 " + HostPathVariable
                    + " 设置为协议宿主 DLL 的完整路径（例如 src\\ThesisAddIn.VSTO\\ProtocolHost\\bin\\Release\\net10.0\\ThesisEngine.ProtocolHost.dll）。";
                return null;
            }

            if (!File.Exists(hostPath))
            {
                error = "找不到引擎宿主文件：" + hostPath + "。请检查 " + HostPathVariable + " 的配置。";
                return null;
            }

            // Full documents may contain hundreds of grouped occurrences. Re-analyzing
            // and annotating that result is more expensive than a single format check.
            return new EngineGateway(new EngineProtocolClient(
                EngineProcessOptions.ForDotnet(hostPath, TimeSpan.FromMinutes(2))));
        }

        public EngineResponse Analyze(string documentPath)
        {
            if (string.IsNullOrWhiteSpace(documentPath))
            {
                return new EngineResponse
                {
                    ProtocolVersion = EngineProtocol.Version,
                    Ok = false,
                    Error = new EngineError { Code = EngineProtocol.InvalidRequest, Message = "当前没有可分析的文档，请先打开一个 .docx 文件。" }
                };
            }

            return _client.Analyze(documentPath);
        }

        public Task<EngineResponse> AnalyzeAsync(string documentPath, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(documentPath))
            {
                return Task.FromResult(new EngineResponse
                {
                    ProtocolVersion = EngineProtocol.Version,
                    Ok = false,
                    Error = new EngineError { Code = EngineProtocol.InvalidRequest, Message = "当前没有可分析的文档，请先打开一个 .docx 文件。" }
                });
            }

            return _client.AnalyzeAsync(documentPath, cancellationToken);
        }

        /// <summary>分析文档并生成用于新建或刷新工作副本的工具批注。</summary>
        public Task<EngineResponse> AnnotateAsync(
            string documentPath,
            string outputPath,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(documentPath) || string.IsNullOrWhiteSpace(outputPath))
            {
                return Task.FromResult(new EngineResponse
                {
                    ProtocolVersion = EngineProtocol.Version,
                    Ok = false,
                    Error = new EngineError
                    {
                        Code = EngineProtocol.InvalidRequest,
                        Message = "生成标注副本时必须提供输入和输出文件路径。"
                    }
                });
            }

            return _client.AnnotateAsync(documentPath, outputPath, cancellationToken);
        }

        /// <summary>生成修复计划；diagnostic 为 null 时计划全部可自动修复项。</summary>
        public EngineResponse Plan(string documentPath, EngineDiagnostic diagnostic)
        {
            return _client.Plan(
                documentPath,
                diagnostic != null ? diagnostic.RuleId : null,
                diagnostic != null ? diagnostic.Location : null,
                GetLocations(diagnostic));
        }

        public Task<EngineResponse> PlanAsync(
            string documentPath,
            EngineDiagnostic diagnostic,
            CancellationToken cancellationToken)
        {
            return _client.PlanAsync(
                documentPath,
                diagnostic != null ? diagnostic.RuleId : null,
                diagnostic != null ? diagnostic.Location : null,
                GetLocations(diagnostic),
                cancellationToken);
        }

        /// <summary>执行修复并写入调用方提供的临时输出路径；diagnostic 为 null 时修复全部可自动项。</summary>
        public EngineResponse Apply(string documentPath, string outputPath, EngineDiagnostic diagnostic)
        {
            return _client.Apply(
                documentPath,
                outputPath,
                diagnostic != null ? diagnostic.RuleId : null,
                diagnostic != null ? diagnostic.Location : null,
                GetLocations(diagnostic));
        }

        public Task<EngineResponse> ApplyAsync(
            string documentPath,
            string outputPath,
            EngineDiagnostic diagnostic,
            CancellationToken cancellationToken)
        {
            return _client.ApplyAsync(
                documentPath,
                outputPath,
                diagnostic != null ? diagnostic.RuleId : null,
                diagnostic != null ? diagnostic.Location : null,
                GetLocations(diagnostic),
                cancellationToken);
        }

        private static System.Collections.Generic.IReadOnlyList<string> GetLocations(EngineDiagnostic diagnostic)
        {
            if (diagnostic == null)
                return null;
            if (diagnostic.Occurrences != null && diagnostic.Occurrences.Count > 0)
                return diagnostic.Occurrences.Select(occurrence => occurrence.Location).Distinct(StringComparer.Ordinal).ToArray();
            return new[] { diagnostic.Location };
        }

        public static string DescribeError(EngineError error)
        {
            if (error == null)
            {
                return "引擎返回了未知错误。";
            }

            switch (error.Code)
            {
                case EngineProtocol.InvalidRequest:
                    return "请求无效：" + error.Message + "（本次未修改文档。）";
                case EngineProtocol.ProtocolVersionMismatch:
                    return "插件与引擎的协议版本不一致，已停止调用：" + error.Message + " 请同步更新插件与引擎宿主。";
                case EngineProtocol.ProtocolError:
                    return "协议通信异常，结果未被采纳：" + error.Message;
                case EngineProtocol.Timeout:
                    return "引擎超时，已终止本次调用：" + error.Message + " 可以稍后重试。";
                case EngineProtocol.EngineExit:
                    return "引擎不可用：" + error.Message;
                case EngineProtocol.EngineError:
                    return "引擎分析失败：" + error.Message + "（本次未写入任何修改。）";
                case EngineProtocol.Cancelled:
                    return "本次调用已取消。";
                default:
                    return "引擎错误（" + error.Code + "）：" + error.Message;
            }
        }
    }
}


