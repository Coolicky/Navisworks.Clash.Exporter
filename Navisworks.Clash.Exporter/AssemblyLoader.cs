using System;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Autodesk.Navisworks.Api.Plugins;

namespace Navisworks.Clash.Exporter

{
    [Plugin("PM.Navisworks.ClashExporter.Loader", "COOL")]
    // ReSharper disable once ClassNeverInstantiated.Global
    public class AssemblyLoader : EventWatcherPlugin
    {
        private static string _thisAssemblyPath;

        public static Assembly ResolveAssemblies(object sender, ResolveEventArgs args)
        {
            var path = Path.GetDirectoryName(_thisAssemblyPath);
            if (path == null) return null;
            var dll = $"{new Regex(",.*").Replace(args.Name, string.Empty)}.dll";
            var file = Path.Combine(path, dll);

            return !File.Exists(file) ? null : Assembly.LoadFrom(file);
        }

        public override void OnLoaded()
        {
            _thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            AppDomain.CurrentDomain.AssemblyResolve += ResolveAssemblies;
        }

        public override void OnUnloading()
        {
            AppDomain.CurrentDomain.AssemblyResolve -= ResolveAssemblies;
        }
    }
}