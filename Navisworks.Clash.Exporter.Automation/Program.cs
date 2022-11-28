using System;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Threading;
using Autodesk.Navisworks.Api.Automation;
using Autodesk.Navisworks.Api.Resolver;
using CommandLine;
using CommandLine.Text;
using Serilog;
using Serilog.Events;

namespace Navisworks.Clash.Exporter.Automation
{
    internal static class Program
    {
        private static readonly string[] AllowedClashGroups =
        {
            "None",
            "Level",
            "GridIntersection",
            "SelectionA",
            "SelectionB",
            "ModelA",
            "ModelB",
            "AssignedTo",
            "ApprovedBy",
            "Status",
            "ItemTypeA",
            "ItemTypeB"
        };

        [STAThread]
        public static void Main(string[] args)
        {
            var runtimeName = Resolver.TryBindToRuntime(RuntimeNames.Any);
            if (string.IsNullOrEmpty(runtimeName))
            {
                throw new Exception("Failed to bind to Navisworks runtime");
            }
            
            var parseResult = Parser.Default.ParseArguments<Options>(args);
            parseResult.WithParsed(RunOptionsAndReturnExitCode)
                .WithNotParsed(errors => ThrowOnParseError(parseResult));
        }

        private static void ThrowOnParseError<T>(ParserResult<T> result)
        {
            var builder = SentenceBuilder.Create();
            var errorMessages = HelpText
                .RenderParsingErrorsTextAsLines(result, builder.FormatError, builder.FormatMutuallyExclusiveSetErrors,
                    1)
                .ToList();
            var lines = HelpText.RenderUsageTextAsLines(result, example => example);
            errorMessages.AddRange(lines);
            foreach (var message in errorMessages)
            {
                Log.Error("{Error}", message);
            }
        }

        private static void RunOptionsAndReturnExitCode(Options options)
        {
            var logFileLocation = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                $@"Navisworks Clash Exporter\Logs\{DateTime.Now:yy-MM-dd}.log");

            if (!string.IsNullOrEmpty(options.LogLocation))
                logFileLocation = Path.Combine(options.LogLocation,
                    $"{DateTime.Now:yy-MM-dd}.log");

            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File(logFileLocation, LogEventLevel.Information, shared: true)
                .CreateLogger();

            Log.Information("Log File Location: {LogFileLocation}", logFileLocation);

            var file = options.NavisworksFile;
            if (!File.Exists(file))
            {
                Log.Error("File {File} does not exist", file);
                return;
            }

            if (!(!Path.GetExtension(file).Equals(".nwf", StringComparison.OrdinalIgnoreCase)
                  || !Path.GetExtension(file).Equals(".nwd", StringComparison.OrdinalIgnoreCase)))
            {
                Log.Error("File {File} is not a Navisworks file", file);
                return;
            }

            var folder = options.ExportFolder;
            if (!Directory.Exists(folder))
            {
                Log.Error("Export folder {Folder} does not exist", folder);
                return;
            }

            if (!HasWriteAccessToFolder(folder))
            {
                Log.Error("Export folder {Folder} is not writeable", folder);
                return;
            }

            if (!AllowedClashGroups.Contains(options.GroupBy))
            {
                Log.Error("GroupBy {GroupBy} is not allowed", options.GroupBy);
                return;
            }

            if (!AllowedClashGroups.Contains(options.ThenBy))
            {
                Log.Error("GroupBy {GroupBy} is not allowed", options.GroupBy);
                return;
            }

            if (options.GroupBy != "None" && options.GroupBy == options.ThenBy)
            {
                Log.Error("Group options {GroupBy} and {ThenBy} are the same", options.GroupBy, options.ThenBy);
                return;
            }

            Export(options);
        }

        private static void Export(Options options)
        {
            NavisworksApplication application = null;

            try
            {
                OpenPipeLineServer();

                Log.Information("Opening Navisworks application");
                application = new NavisworksApplication();
                if (options.HideUI)
                    application.DisableProgress();
                application.Visible = !options.HideUI;
                Log.Information("Opening Navisworks File: {File}", options.NavisworksFile);
                application.OpenFile(options.NavisworksFile);
                var parameters = options.WriteToArray();
                Log.Information("Export Parameters");
                foreach (var parameter in parameters)
                {
                    Log.Information("{Parameter}", parameter);
                }

                Log.Information("Executing Clash Export");
                var result =
                    application.ExecuteAddInPlugin("Navisworks.Clash.Exporter.ExportAutomatic.COOL", parameters);


                if (result == 1)
                {
                    Log.Information("Clash Export Successful");
                    if (!options.SkipFileSave)
                    {
                        Log.Information("Saving File...");
                        application.SaveFile(options.NavisworksFile);
                    }
                }
                else
                {
                    Log.Error("Export Clash Failed");
                }

                application.EnableProgress();
            }
            catch (Exception e)
            {
                Log.Error("Error:{Message}\n{Stack}", e.Message, e.StackTrace);
            }
            finally
            {
                Log.Information("Closing all Navisworks Resources");
                application?.Dispose();
            }
        }

        private static void OpenPipeLineServer()
        {
            var thread = new Thread(() =>
            {
                using (var pipeStream = new NamedPipeServerStream("Clash Exporter", PipeDirection.InOut))
                {
                    Log.Information("[Server] Pipe Created");
                    pipeStream.WaitForConnection();

                    Log.Information("[Server] Pipe connection established");

                    using (var sr = new StreamReader(pipeStream))
                    {
                        string message;
                        while ((message = sr.ReadLine()) != null)
                        {
                            Log.Information("{Message}", message);
                        }
                    }

                    Log.Information("Pipe Disconnected");
                }
            });
            thread.Start();
        }

        private static bool HasWriteAccessToFolder(string folderPath)
        {
            try
            {
                Directory.GetAccessControl(folderPath);
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
        }
    }
}