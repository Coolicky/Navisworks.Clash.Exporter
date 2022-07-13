using System;
using System.IO;
using Serilog.Core;
using Serilog.Events;

namespace Navisworks.Clash.Exporter.Logger
{
    public class PipeSink : ILogEventSink
    {
        private readonly IFormatProvider _formatProvider;
        private readonly StreamWriter _writer;

        public PipeSink(IFormatProvider formatProvider, StreamWriter writer)
        {
            _formatProvider = formatProvider;
            _writer = writer;
        }
        public void Emit(LogEvent logEvent)
        {
            var message = logEvent.RenderMessage(_formatProvider);
            _writer.WriteLine(message);
            _writer.Flush();
        }
    }
}