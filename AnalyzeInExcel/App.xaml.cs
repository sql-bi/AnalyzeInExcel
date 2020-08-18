using CommandLine;
using System;
using System.Data;
using System.Linq;
using System.Windows;
using System.Diagnostics;
using AutoUpdaterDotNET;
using System.Windows.Input;
using System.Threading.Tasks;

namespace AnalyzeInExcel
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public Options AppOptions;
        const string MSOLAP_DRIVER_URL = @"https://go.microsoft.com/fwlink/?LinkId=746283";

        const string EV_MSOLAP_NOTFOUND = "MSOLAP driver not found";
        const string EV_MSOLAP_SETUP = "Requested MSOLAP driver setup";
        const string EV_MODEL_NOT_AVAILABLE = "Model not available";
        const string EV_EXCEL_NOT_AVAILABLE = "Excel not available";
        const string EV_CONFIGURATION_INCOMPLETE = "Configuration incomplete";
        const string EV_RUNEXCEL = "Run Excel";

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
        public MessageBoxResult ShowMessageQuestion(string message)
        {
            return MessageBox.Show(message, "Analyze in Excel for Power BI Desktop",MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes);
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

            try
            {
                string serverName = ((App)Application.Current).AppOptions?.Server;
                string databaseName = ((App)Application.Current).AppOptions?.Database;
                bool goodMsOlapDriver = ModelHelper.HasMsOlapDriver();
                if (!goodMsOlapDriver)
                {
                    th.TrackEvent(EV_MSOLAP_NOTFOUND);
                    if (ShowMessageQuestion($"Excel needs a component called MSOLAP driver to connect to Power BI. The MSOLAP driver might be missing or not updated on this device. Therefore, Excel might not connect to Power BI. You can install the updated Microsoft MSOLAP driver from this link: {MSOLAP_DRIVER_URL} \n\nClick YES if you want to download the updated MSOLAP driver and install it.\nClick NO to continue without any update.") == MessageBoxResult.Yes)
                    {
                        try
                        {
                            th.TrackEvent(EV_MSOLAP_SETUP);
                            UpdateMsOlapDriver();
                        }
                        catch (Exception ex)
                        {
                            // Send any exception to Telemetry
                            th.TrackException(ex);
                            ShowMessage($"Error running MSOLAP update: {ex.Message}");
                        }
                    }
                }
                // We use the default "Model" string if the MSOLAP driver is not available - if this happens, the previous warning helps understanding possible issues
                string cubeName = goodMsOlapDriver ? ModelHelper.GetModelName(serverName, databaseName, th) : "Model"; 
                if (serverName != null && databaseName != null)
                {
                    try
                    {
                        if (string.IsNullOrEmpty(cubeName))
                        {
                            ShowMessage("Power BI has an empty model or it is connected to an unkonwn external dataset. You cannot connect Excel.");
                            th.TrackEvent(EV_MODEL_NOT_AVAILABLE);
                        }
                        else if (ExcelHelper.IsExcelAvailable())
                        {
                            // TODO: Manage options requested
                            if (OptionsRequested)
                            {
                                // TODO request action / configuration to users
                            }

                            var splashScreen = new SplashLoading();
                            try
                            {
                                this.MainWindow = splashScreen;
                                splashScreen.Show();

                                bool excelStarted = ExcelHelper.CreateInstanceWithPivotTable(serverName, databaseName, cubeName, (ex) => th.TrackException(ex));
                                if (excelStarted)
                                {
                                    th.TrackEvent(EV_RUNEXCEL, "RunType", "Interop");
                                }
                                else {
                                    RunExcelProcess(serverName, databaseName, cubeName);
                                    th.TrackEvent(EV_RUNEXCEL, "RunType", "ODC File");
                                }
                            }
                            finally
                            {
                                splashScreen.Close();
                            }
                        }
                        else
                        {
                            ShowMessage("Excel is not available. Please check whether Excel is correctly installed.");
                            th.TrackEvent(EV_EXCEL_NOT_AVAILABLE);
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
                    th.TrackEvent(EV_CONFIGURATION_INCOMPLETE);
                    th.Flush();
                }

                // Check updates asynchronously when there is an error, while displaying the diagnostic message
                CheckUpdates(false);

                OpenDiagnosticWindow(serverName, databaseName);
            }
            catch (Exception ex)
            {
                // Send any exception to Telemetry
                th.TrackException(ex);
                th.Flush();
                throw;
            }

            void UpdateMsOlapDriver()
            {
                var p = new Process
                {
                    StartInfo = new ProcessStartInfo(MSOLAP_DRIVER_URL)
                    {
                        UseShellExecute = true
                    }
                };
                p.Start();
            }
        }
        private void OpenDiagnosticWindow(string serverName, string databaseName)
        {
            //// start application window
            MainWindow mw = new MainWindow();
            this.MainWindow = mw;
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

        private void RunExcelProcess(string serverName, string databaseName, string cubeName)
        {
            // Create ODC file
            var fileName = OdcHelper.CreateOdcFile(serverName, databaseName, cubeName);

            // Open ODC file
            var p = new Process
            {
                StartInfo = new ProcessStartInfo(fileName)
                {
                    UseShellExecute = true
                }
            };
            p.Start();
        }
    }
}
