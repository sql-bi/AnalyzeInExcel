using CommandLine;

namespace AnalyzeInExcel
{
    public class Options
    {
        [Option('d', "database", HelpText = "Database name")]
        public string Database { get; set; }

        [Option('s', "server", HelpText = "Server name")]
        public string Server { get; set; }

        [Option('n', "telemetry", HelpText = "Enable Telemetry")]
        public bool Telemetry { get; set; }

        [Option('x', "experiment", HelpText = "Experiment feature")]
        public bool Experiment { get; set; }
    }
}
