using Microsoft.ApplicationInsights.DataContracts;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.Remoting.Contexts;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Reflection;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;

namespace ExternalToolsInstaller
{
    [RunInstaller(true)]
    public partial class ExternalToolsInstaller : System.Configuration.Install.Installer
    {
        /// <summary>
        /// Name of external tool executable
        /// </summary>
        const string EXTERNALTOOLS_EXENAME = @"AnalyzeInExcel.exe";

        /// <summary>
        /// Name of external tool configuration file
        /// </summary>
        const string EXTERNALTOOLS_CONFIGFILENAME = @"analyzeinexcel.pbitool.json";

        /// <summary>
        /// The "assemblypath" setup argument corresponds to the path of the DLL invoked by the installer 
        /// we assume it is the same folder where we installed the AnalyzeInExcel.exe file
        /// </summary>
        const string SETUP_ASSEMBLYPATH = "assemblypath";

        /// <summary>
        /// The telemetry setup argument correspond to the radio button selection
        /// to enable or disable telemetry (1=enabled, 0=disabled)
        /// </summary>
        const string SETUP_TELEMETRY = "telemetry";

        /// <summary>
        /// The product version setup argument correspond to the version of the product installed
        /// </summary>
        const string SETUP_PRODUCT_VERSION = "version";

        /// <summary>
        /// Command line argument to enable telemetry in external tool
        /// Add this command line argument if the telemetry is enabled
        /// </summary>
        const string TELEMETRY_ARGUMENT = "--telemetry";

        public ExternalToolsInstaller()
        {
            InitializeComponent();

            PathConfiguration = new PathExternalToolConfiguration(EXTERNALTOOLS_CONFIGFILENAME);
        }

        private PathExternalToolConfiguration PathConfiguration { get; }

        private ExternalToolConfiguration ReadExternalToolConfiguration()
        {
            string config = System.IO.File.ReadAllText(PathConfiguration.FullPath);
            return JsonConvert.DeserializeObject(config, typeof(ExternalToolConfiguration)) as ExternalToolConfiguration;
        }

        private void WriteExternalToolConfiguration(ExternalToolConfiguration configuration)
        {
            string updatedConfig = JsonConvert.SerializeObject(configuration, Formatting.Indented);
            System.IO.File.WriteAllText(PathConfiguration.FullPath, updatedConfig);
        }

        private string GetInstallerProductVersion()
        {
            if (Context.Parameters.ContainsKey(SETUP_PRODUCT_VERSION))
            {
                return Context.Parameters[SETUP_PRODUCT_VERSION];
            }
            else
            {
                return null;
            }
        }

        private bool IsSetupTelemetryEnabled()
        {
            if (Context.Parameters.ContainsKey(SETUP_TELEMETRY))
            {
                string telemetryValue = Context.Parameters[SETUP_TELEMETRY] ?? string.Empty;
                return (telemetryValue.Trim() != "0");
            }
            else
            {
                // In case of missing argument enable telemetry to further investigate
                return true;
            }
        }

        private TelemetryClient GetTelemetryClient( bool telemetryEnabled )
        {
            TelemetryConfiguration telemetryConfiguration = TelemetryConfiguration.CreateDefault();
            telemetryConfiguration.InstrumentationKey = "60ab83db-108f-45ee-b537-d70dc47d3193";
            telemetryConfiguration.DisableTelemetry = (telemetryEnabled == false);
            TelemetryClient tc = new Microsoft.ApplicationInsights.TelemetryClient(telemetryConfiguration);
            return tc;
        }

        private void ExternalToolsInstaller_AfterInstall(object sender, InstallEventArgs e)
        {
            bool telemetryEnabled = IsSetupTelemetryEnabled();
            var ev = new EventTelemetry();
            ev.Name = "Install";

            // Fix version installed
            ExternalToolConfiguration externalToolConfiguration = ReadExternalToolConfiguration();
            externalToolConfiguration.version = GetInstallerProductVersion() ?? externalToolConfiguration.version;
            ev.Properties["Version"] = externalToolConfiguration.version;

            // Fix executable in configuration
            string assemblyPath;
            if (Context.Parameters.ContainsKey(SETUP_ASSEMBLYPATH))
            {
                assemblyPath = Context.Parameters[SETUP_ASSEMBLYPATH];
            }
            else
            {
                Context.LogMessage($"The {"ASSEMBLYPATH"} property is not available, assume directory of installer DLL instead.");
                assemblyPath = Assembly.GetExecutingAssembly().Location;
            }
            string externalToolsDirectory = System.IO.Path.GetDirectoryName(assemblyPath);
            string externalToolsExe = System.IO.Path.Combine(externalToolsDirectory, EXTERNALTOOLS_EXENAME);
            externalToolConfiguration.path = externalToolsExe;

            // Remove existing telemetry option
            string telemetryOption = $" {TELEMETRY_ARGUMENT}";
            externalToolConfiguration.arguments = externalToolConfiguration.arguments.Replace(telemetryOption, "");
            externalToolConfiguration.arguments = externalToolConfiguration.arguments.Replace(telemetryOption.Trim(), "");

            // Add telemetry argument if telemetry is enabled during setup
            if (telemetryEnabled)
            {
                externalToolConfiguration.arguments += telemetryOption;
                Context.LogMessage($"Telemetry enabled");
            }

            // Initialize Telemetry
            TelemetryClient tc = GetTelemetryClient(telemetryEnabled);
            try
            {
                // Update external tool configuration
                WriteExternalToolConfiguration(externalToolConfiguration);

                // Send telemetry
                tc.TrackEvent(ev);
                tc.Flush();
            }
            catch (Exception ex)
            {
                // In case of error, send exception to Telemetry
                tc.TrackException(ex);
                tc.Flush();
                throw;
            }
        }

        private void ExternalToolsInstaller_AfterUninstall(object sender, InstallEventArgs e)
        {
            // Read version installed
            ExternalToolConfiguration externalToolConfiguration = ReadExternalToolConfiguration();
            string versionInstalled = GetInstallerProductVersion() ?? externalToolConfiguration?.version ?? "unknown";
            
            // Enable telemetry if it was enabled in configuration file or if there is no configuration
            bool telemetryEnabled = externalToolConfiguration?.arguments?.Contains(TELEMETRY_ARGUMENT) ?? true;
            var ev = new EventTelemetry();
            ev.Name = "Uninstall";
            ev.Properties["Version"] = versionInstalled;

            // Initialize Telemetry
            TelemetryClient tc = GetTelemetryClient(telemetryEnabled);
            try
            {
                // Send telemetry
                tc.TrackEvent(ev);
                tc.Flush();
            }
            catch (Exception ex)
            {
                // In case of error, send exception to Telemetry
                tc.TrackException(ex);
                tc.Flush();
                throw;
            }
        }
    }
}
