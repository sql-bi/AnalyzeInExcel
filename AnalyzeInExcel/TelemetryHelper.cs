using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DataContracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnalyzeInExcel
{
    public class TelemetryHelper
    {
        protected TelemetryClient TC { get; }

        private string Version { get; }

        public TelemetryHelper( bool disableTelemetry = false )
        {
            Microsoft.ApplicationInsights.Extensibility.TelemetryConfiguration configuration = Microsoft.ApplicationInsights.Extensibility.TelemetryConfiguration.CreateDefault();
            configuration.InstrumentationKey = "60ab83db-108f-45ee-b537-d70dc47d3193";
            configuration.DisableTelemetry = disableTelemetry;
            TC = new TelemetryClient(configuration);

            var assemblyVersion = this.GetType().Assembly.GetName().Version;
            Version = $"{assemblyVersion.Major}.{assemblyVersion.Minor}.{assemblyVersion.Build}";
        }

        protected virtual EventTelemetry CreateEvent( string eventName )
        {
            var ev = new EventTelemetry
            {
                Name = eventName
            };
            ev.Properties["Version"] = Version;
            return ev;
        }

        public void TrackEvent( string eventName )
        {
            TC.TrackEvent(CreateEvent(eventName));
        }

        public void TrackException(Exception exception)
        {
            TC.TrackException(exception);
        }

        public void TrackException(ExceptionTelemetry exceptionTelemetry)
        {
            TC.TrackException(exceptionTelemetry);
        }

        public void Flush()
        {
            TC.Flush();
        }
    }
}
