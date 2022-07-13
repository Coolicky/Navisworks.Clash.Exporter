using System;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Security.Principal;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using CommandLine;
using CommandLine.Text;
using Navisworks.Clash.Exporter.Extensions;
using Navisworks.Clash.Exporter.Methods;
using Navisworks.Clash.Exporter.Models;
using Serilog;

namespace Navisworks.Clash.Exporter
{
    [Plugin("Navisworks.Clash.Exporter.ExportAutomatic", "COOL")]
    [AddInPlugin(AddInLocation.None, LoadForCanExecute = true,
        CallCanExecute = CallCanExecute.DocumentNotClear)]
    public class ExportAutomatic : AddInPlugin
    {    
        private static NamedPipeClientStream _pipeClient;
        private static StreamWriter _streamWriter;
        private ILogger _log;

        public override int Execute(params string[] parameters)
        {
            AppDomain.CurrentDomain.AssemblyResolve += AssemblyLoader.ResolveAssemblies;
            try
            {
                _pipeClient = new NamedPipeClientStream(".","Clash", PipeDirection.InOut, PipeOptions.None, TokenImpersonationLevel.Anonymous);
                _pipeClient.Connect();
                _streamWriter = new StreamWriter(_pipeClient);
                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Debug()
                    .WriteTo.PipeSink(_streamWriter)
                    .CreateLogger();
                _log = Log.ForContext<ExportAutomatic>();
                _log.Information("Application: {Version} Started", Application.Version.RuntimeProductName);
                _log.Information("PipeLine Connected");
            }
            catch
            {
                // ignored
            }

            try
            {
                _log.Information("Parsing Command Line Arguments");
                var parserResult = Parser.Default.ParseArguments<CommandLineOptions>(parameters);
                if (parserResult.Tag == ParserResultType.Parsed)
                {
                    var options = ((Parsed<CommandLineOptions>)parserResult).Value;
                    NavisworksExtensions.Imperial = options.Imperial;

                    _log.Information("Successfully Parsed proceeding to export...");
                    return ClashExport.ExportClashes(options);
                }

                _log.Error("Parsing command line arguments failed");
                ThrowOnParseError(parserResult);
                return 0;
            }
            catch (Exception e)
            {
                _log.Error("Error while exporting clashes: {Error} \n {Stack}", e.Message, e.StackTrace);
                return 0;
            }
        }

        private static void ThrowOnParseError<T>(ParserResult<T> result)
        {
            var builder = SentenceBuilder.Create();
            var errorMessages = HelpText.RenderParsingErrorsTextAsLines(result, builder.FormatError, builder.FormatMutuallyExclusiveSetErrors, 1)
                .ToList();
            var lines = HelpText.RenderUsageTextAsLines(result, example => example);
            errorMessages.AddRange(lines);
            foreach (var message in errorMessages)
            {
                Log.Error("{Error}", message);
            }
        }
    }
}