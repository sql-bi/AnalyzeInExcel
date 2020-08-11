using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnalyzeInExcel
{
    public class ModelHelper
    {
        const string MSOLAP_DRIVER_NAME = "MSOLAP"; // We get the latest available. We currently don't enforce MSOLAP.8

        internal enum AsInstanceType
        {
            Other,
            AsAzure,
            PbiDedicated,
            PbiPremium,
            PbiDataset
        }

        private static bool IsProtocolSchemeInstance(string dataSourceUri, string protocolScheme)
        {
            return dataSourceUri?.StartsWith(protocolScheme, StringComparison.InvariantCultureIgnoreCase) ?? false;
        }
        internal static AsInstanceType GetAsInstanceType(string dataSourceUri)
        {
            if (IsProtocolSchemeInstance(dataSourceUri, "asazure://"))
            {
                return AsInstanceType.AsAzure;
            }
            if (IsProtocolSchemeInstance(dataSourceUri, "pbidedicated://"))
            {
                return AsInstanceType.PbiDedicated;
            }
            if (IsProtocolSchemeInstance(dataSourceUri, "powerbi://"))
            {
                return AsInstanceType.PbiPremium;
            }
            if (IsProtocolSchemeInstance(dataSourceUri, "pbiazure://"))
            {
                return AsInstanceType.PbiDataset;
            }
            return AsInstanceType.Other;
        }

        /// <summary>
        /// Check whether the machine has the MSOLAP driver installed
        /// The driver could be available in 32-bit and not in 64-bit
        /// This would work in Excel, but it is better to try get the driver in that case
        /// </summary>
        /// <returns></returns>
        public static bool HasMsOlapDriver()
        {
            const string SOURCES_NAME = "SOURCES_NAME";

            var oleEnum = new OleDbEnumerator();
            var elems = oleEnum.GetElements();
            if (elems != null && elems.Rows != null)
                foreach (System.Data.DataRow row in elems.Rows)
                    if (!row.IsNull(SOURCES_NAME) && row[SOURCES_NAME] is string)
                        if (row[SOURCES_NAME].ToString() == MSOLAP_DRIVER_NAME )
                            return true;
            return false;
        }

        /// <summary>
        /// Returns the OLE DB connection string based on serverName and databaseName
        /// </summary>
        /// <param name="serverName"></param>
        /// <param name="databaseName"></param>
        /// <returns>Connection string for OLE DB</returns>
        public static string GetOleDbConnectionString( string serverName, string databaseName )
        {
            // Choose the proper connection string
            string connectionString;
            switch (GetAsInstanceType(serverName))
            {
                case AsInstanceType.PbiDataset: // Integrated Security=ClaimsToken;
                    connectionString = $"Provider={MSOLAP_DRIVER_NAME};Persist Security Info=True;Initial Catalog=sobe_wowvirtualserver-{databaseName};Data Source={serverName};MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error;Identity Provider=https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, 929d0ec0-7a41-4b1e-bc7c-b754a28bddcc;Update Isolation Level=2";
                    break;
                case AsInstanceType.PbiDedicated:
                case AsInstanceType.PbiPremium:
                case AsInstanceType.AsAzure:
                    connectionString = $"Provider={MSOLAP_DRIVER_NAME};Persist Security Info=True;Data Source={serverName};Update Isolation Level=2;Initial Catalog={databaseName}";
                    break;
                case AsInstanceType.Other:
                default:
                    connectionString = $"Provider={MSOLAP_DRIVER_NAME};Integrated Security=SSPI;Persist Security Info=True;Data Source={serverName};Update Isolation Level=2;Initial Catalog={databaseName}";
                    break;
            }
            return connectionString;
        }

        /// <summary>
        /// Check whether the connection has a data model
        /// </summary>
        /// <param name="serverName">Server name</param>
        /// <param name="databaseName">Database name</param>
        /// <returns>true it the connection has a data model</returns>
        public static string GetModelName(string serverName, string databaseName, TelemetryHelper th)
        {
            string result = null;
            string connectionString = GetOleDbConnectionString(serverName, databaseName);
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    using (OleDbCommand command = new OleDbCommand("select CUBE_NAME from $SYSTEM.MDSCHEMA_CUBES", connection))
                    {
                        connection.Open();
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                result = reader.GetString(0);
                            }
                            reader.Close();
                        }
                    }
                }
            }
            catch (System.Data.Common.DbException ex)
            {
                // Ignore DbException and return a null model name

                // Send exception to Telemetry for further investigation
                th.TrackException(ex);
                th.Flush();
            }
            return result;
        }
    }
}
