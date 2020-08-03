using CommandLine;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;
using System.Diagnostics;

namespace AnalyzeInExcel
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public Options AppOptions;
        protected override void OnStartup(StartupEventArgs e)
        {
            // Read configuration
            var result = Parser.Default.ParseArguments<Options>(e.Args)
                .WithNotParsed(errors => MessageBox.Show("Invalid configuration, check the pbitool.json file\n" + String.Join(";", from err in errors select err.ToString())))
                .WithParsed(options => AppOptions = options);

            Microsoft.ApplicationInsights.Extensibility.TelemetryConfiguration configuration = Microsoft.ApplicationInsights.Extensibility.TelemetryConfiguration.CreateDefault();
            configuration.InstrumentationKey = "60ab83db-108f-45ee-b537-d70dc47d3193";
            configuration.DisableTelemetry = (AppOptions.Telemetry == false);
            Microsoft.ApplicationInsights.TelemetryClient tc = new Microsoft.ApplicationInsights.TelemetryClient(configuration);

            string serverName = ((App)Application.Current).AppOptions?.Server;
            string databaseName = ((App)Application.Current).AppOptions?.Database;
            string cubeName = "Model"; // TODO get this from the model?
            if (serverName != null && databaseName != null)
            {
                try
                {
                    // Create ODC file
                    OdcHelper.CreateOdcFile(serverName, databaseName, cubeName);
                    var fileName = OdcHelper.OdcFilePath();

                    // Open ODC file
                    var p = new Process();
                    p.StartInfo = new ProcessStartInfo(fileName)
                    {
                        UseShellExecute = true
                    };
                    tc.TrackEvent("Run Excel");
                    p.Start();
                    tc.Flush();
                    this.Shutdown(0);
                }
                catch (Exception ex)
                {
                    tc.TrackException(ex);
                    tc.Flush();
                    MessageBox.Show("Error launching Excel: " + ex.Message);
                }
            }
            else
            {
                tc.TrackEvent("Configuration incomplete");
                tc.Flush();
            }

            //// start application window
            MainWindow mw = new MainWindow();
            mw.diagnosticInfo.Content = $@"Current configuration
Server={serverName ?? "(blank)"}
Database={databaseName ?? "(blank)"}

Check that the file 
{Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86)}\Microsoft Shared\Power BI Desktop\External Tools\analyzeinexcel.pbitool.json 
includes the following argument:
  ""arguments"": ""--telemetry --server =\""%server%\"" --database=\""%database%\""
";
            mw.Show();
        }
    }
}
