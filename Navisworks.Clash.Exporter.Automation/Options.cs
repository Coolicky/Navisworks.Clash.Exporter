using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using CommandLine;

namespace Navisworks.Clash.Exporter.Automation
{
    [SuppressMessage("ReSharper", "UnusedAutoPropertyAccessor.Global")]
    public class Options
    {
        [Option('n', "navisworks", Required = true, HelpText = "Navisworks (.nwd/.nwf) file")]
        public string NavisworksFile { get; set; }

        [Option('f', "exportFolder", Required = true, HelpText = "Export folder")]
        public string ExportFolder { get; set; }
        
        [Option('l', "logLocation", Required = true, HelpText = "Export folder")]
        public string LogLocation { get; set; }

        [Option("imperial", HelpText = "Whether to use imperial units (feet)")]
        public bool Imperial { get; set; }

        [Option("groupBy", Required = false, Default = "None", HelpText =
            "What to Group Clashes By. Default is 'None'\n" +
            "Options are 'Level', 'GridIntersection', 'SelectionA', 'SelectionB', 'ModelA', 'ModelB', 'AssignedTo', 'ApprovedBy', 'Status', 'ItemTypeA', 'ItemTypeB'")]
        public string GroupBy { get; set; }

        [Option("thenBy", Required = false, Default = "None", HelpText =
            "What to Group Clashes By after. Default is 'None'\n" +
            "Options are 'Level', 'GridIntersection', 'SelectionA', 'SelectionB', 'ModelA', 'ModelB', 'AssignedTo', 'ApprovedBy', 'Status', 'ItemTypeA', 'ItemTypeB'")]
        public string ThenBy { get; set; }

        [Option("keepGroups", Required = false, HelpText = "Whether to Keep Existing Groups")]
        public bool KeepGroups { get; set; }

        [Option("skipRefresh", Required = false, HelpText = "Whether to Run Clashes")]
        public bool SkipRefresh { get; set; }

        [Option("savePrevious", HelpText = "Saves previous export to a superseded folder")]
        public bool SavePrevious { get; set; }

        [Option("skipFileSave", HelpText = "Does not save Navisworks File. Useful for testing")]
        public bool SkipFileSave { get; set; }

        [Option("hideUI", HelpText = "Hides Navisworks User Interface")]
        public bool HideUI { get; set; }


        public string[] WriteToArray()
        {
            var options = new List<string>();

            options.Add("-n");
            options.Add(NavisworksFile);
            options.Add("-f");
            options.Add(ExportFolder);
            options.Add("--groupBy");
            options.Add(GroupBy);
            options.Add("--thenBy");
            options.Add(ThenBy);
            if (KeepGroups) options.Add("--keepGroups");
            if (SkipRefresh) options.Add("--runClash");
            if (SavePrevious) options.Add("--savePrevious");
            if (Imperial) options.Add("--imperial");
            return options.ToArray();
        }
    }
}