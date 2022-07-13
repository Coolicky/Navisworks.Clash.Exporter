using System;
using System.IO;
using Navisworks.Clash.Exporter.Logger;
using Serilog;
using Serilog.Configuration;

namespace Navisworks.Clash.Exporter.Extensions
{
    public static class SerilogExtensions
    {
        public static LoggerConfiguration PipeSink(
            this LoggerSinkConfiguration loggerConfig,
            StreamWriter streamWriter,
            IFormatProvider formatProvider = null)
        {
            return loggerConfig.Sink(new PipeSink(formatProvider, streamWriter));
        }
    }
}