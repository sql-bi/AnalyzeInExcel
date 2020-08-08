using CommandLine;
using System;
using System.Data;
using System.Linq;
using System.Windows;
using System.Diagnostics;
using AutoUpdaterDotNET;
using System.Windows.Input;

namespace AnalyzeInExcel
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public Options AppOptions;

        public void CheckUpdates(bool synchronous = true)
        {
            AutoUpdater.HttpUserAgent = "AutoUpdater";
            AutoUpdater.ShowSkipButton = true;
            AutoUpdater.ShowRemindLaterButton = true;
            AutoUpdater.LetUserSelectRemindLater = true;
            AutoUpdater.Synchronous = synchronous;
            // The random number guarantees that the web client cache is not used (it is applied often even though internally the policy is disabled)
            AutoUpdater.Start($@"https://cdn.sqlbi.com/updates/AnalyzeInExcelAutoUpdater.xml?random={new Random().Next()}");
        }

        /// <summary>
        /// THe user requested an options window instead of running the default action
        /// </summary>
        public bool OptionsRequested { get; private set; }

        /// <summary>
        /// Check conditions to open Options window instead of directly opening Excel 
        /// </summary>
        protected void InitializeOptionRequested()
        {
            // Pressing CTRL (left or right) when the external tool is launched activates the option window
            OptionsRequested = System.Windows.Input.Keyboard.IsKeyDown(Key.LeftCtrl) || System.Windows.Input.Keyboard.IsKeyDown(Key.LeftCtrl);
        }

        public MessageBoxResult ShowMessage( string message )
        {
            return MessageBox.Show(message, "Analyze in Excel for Power BI Desktop");
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            // Store request to access options
            InitializeOptionRequested();

            // Read configuration
            var result = Parser.Default.ParseArguments<Options>(e.Args)
                .WithNotParsed(errors => ShowMessage("Invalid configuration, check the pbitool.json file\n" + String.Join(";", from err in errors select err.ToString())))
                .WithParsed(options => AppOptions = options);

            bool disableTelemetry = (AppOptions.Telemetry == false);
            TelemetryHelper th = new TelemetryHelper(disableTelemetry);

            string serverName = ((App)Application.Current).AppOptions?.Server;
            string databaseName = ((App)Application.Current).AppOptions?.Database;
            string cubeName = ModelHelper.GetModelName(serverName, databaseName);
            if (serverName != null && databaseName != null)
            {
                try
                {
                    if (string.IsNullOrEmpty(cubeName))
                    {
                        ShowMessage("Power BI has an empty model or it is connected to an unkonwn external dataset. You cannot connect Excel.");
                        th.TrackEvent("Model not available");
                    }
                    else if (ExcelHelper.IsExcelAvailable())
                    {
                        // TODO: Manage options requested
                        if (OptionsRequested)
                        {
                            // TODO request action / configuration to users
                        }

                        // Create ODC file
                        OdcHelper.CreateOdcFile(serverName, databaseName, cubeName);
                        var fileName = OdcHelper.OdcFilePath();

                        // Open ODC file
                        var p = new Process
                        {
                            StartInfo = new ProcessStartInfo(fileName)
                            {
                                UseShellExecute = true
                            }
                        };
                        th.TrackEvent("Run Excel");
                        p.Start();
                    }
                    else
                    {
                        ShowMessage("Excel is not available. Please check whether Excel is correctly installed.");
                        th.TrackEvent("Excel not available");
                    }
                    th.Flush();
                    // Check updates synchronously when Excel starts, no wait for Excel
                    CheckUpdates(true);
                    this.Shutdown(0);
                }
                catch (Exception ex)
                {
                    th.TrackException(ex);
                    th.Flush();
                    ShowMessage("Error launching Excel: " + ex.Message);
                }
            }
            else
            {
                th.TrackEvent("Configuration incomplete");
                th.Flush();
            }

            // Check updates asynchronously when there is an error, while displaying the diagnostic message
            CheckUpdates(false);

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
