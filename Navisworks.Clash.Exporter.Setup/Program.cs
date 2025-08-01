﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;
using Navisworks.Clash.Exporter.Setup.Manifest;
using WixSharp;

namespace Navisworks.Clash.Exporter.Setup
{
    internal class Program
    {
        const string Guid = "89C569FF-EC52-4394-BEEB-626A7E239752";
        const string PluginName = "Navisworks Clash Exporter";
        const string DllName = "Navisworks.Clash.Exporter";

        private const string ProjectLocation = @"..\Navisworks.Clash.Exporter";
        private const string AutomationProjectLocation = @"..\Navisworks.Clash.Exporter.Automation";

        private static readonly Dictionary<int, string> Versions = new Dictionary<int, string>
        {
            { 2018, "net46" },
            { 2019, "net47" },
            { 2020, "net47" },
            { 2021, "net47" },
            { 2022, "net47" },
            { 2023, "net48" },
            { 2024, "net48" },
            { 2025, "net48" },
        };

        public static void Main()
        {
            var folders = Versions.ToDictionary(
                r => r.Key.ToString(),
                r => $@"{ProjectLocation}\bin\x64\Release_{r.Key}\{r.Value}");
            var automationFolders = Versions.ToDictionary(
                r => r.Key.ToString(),
                r => $@"{AutomationProjectLocation}\bin\x64\Release_{r.Key}\{r.Value}");

            AutoElements.DisableAutoKeyPath = true;
            var feature = new Feature(PluginName, true, false);

            CreateManifest();

            var directories = CreateDirectories(feature, folders);
            var automationDirectories = CreateDirectories(feature, automationFolders);
            var dir = new Dir(feature, $@"%AppData%/Autodesk/ApplicationPlugins/{DllName}.bundle",
                new File(feature, "./PackageContents.xml"),
                new Dir(feature, "Contents")
                {
                    Dirs = directories
                },
                new Dir(feature, "Automation")
                {
                    Dirs = automationDirectories
                });

            var project = new Project(PluginName, dir)
            {
                Name = PluginName,
                Description = PluginName,
                OutFileName = PluginName,
                OutDir = "output",
                Platform = Platform.x64,
                UI = WUI.WixUI_Minimal,
                Version = GetVersion(),
                Scope = InstallScope.perUser,
                MajorUpgrade = MajorUpgrade.Default,
                GUID = new Guid(Guid),
                LicenceFile = "./Resources/GNU.rtf",
                ControlPanelInfo =
                {
                    ProductIcon = "./Resources/roamer.exe.ico",
                },
                BannerImage = "./Resources/Banner.bmp",
                BackgroundImage = "./Resources/Main.bmp",
            };
            Compiler.BuildMsi(project);
        }

        private static void CreateManifest()
        {
            var components = Versions.Select(r => new Components
            {
                Description = r.Key.ToString(),
                RuntimeRequirements = new RuntimeRequirements()
                {
                    OS = "Win64",
                    Platform = "NAVMAN",
                    SeriesMin = $"Nw{r.Key - 2003}",
                    SeriesMax = $"Nw{r.Key - 2003}"
                },
                ComponentEntry = new ComponentEntry
                {
                    AppName = $"{PluginName} for Autodesk® Navisworks® {r.Key}",
                    AppType = "ManagedPlugin",
                    ModuleName = $"./Contents/{r.Key}/{DllName}.dll"
                }
            }).ToArray();

            var manifest = new ApplicationPackage
            {
                SchemaVersion = 1.0,
                AutodeskProduct = "Naviswork",
                Name = PluginName,
                Description = PluginName,
                AppVersion = $"{GetVersion().Major}.{GetVersion().Minor}",
                ProductCode = System.Guid.NewGuid().ToString(),
                UpgradeCode = Guid,
                CompanyDetails = new CompanyDetails()
                {
                    Name = "Coolicky, https://github.com/Coolicky",
                },
                Components = components
            };
            var serializer = new XmlSerializer(typeof(ApplicationPackage));
            using (var writer = new System.IO.StreamWriter("./PackageContents.xml"))
            {
                serializer.Serialize(writer, manifest);
            }
        }

        private static Dir[] CreateDirectories(Feature feature, Dictionary<string, string> folders)
        {
            var dirs = new List<Dir>();
            foreach (var folder in folders)
            {
                var dir = new Dir(folder.Key,
                    new Files(feature,
                        $@"{folder.Value}\*.*"));
                dirs.Add(dir);
            }

            return dirs.ToArray();
        }

        private static Version GetVersion()
        {
            const int majorVersion = 0;
            const int minorVersion = 14;
            var version = $"{majorVersion}.{minorVersion}.{0}.{0}";
            return new Version(version);
        }
    }
}