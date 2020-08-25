using System;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace AnalyzeInExcel
{
    public class ExcelHelper
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        /// <summary>
        /// Check whether Excel is installed
        /// </summary>
        /// <returns>true if Excel is installed</returns>
        public static bool IsExcelAvailable()
        {
            var type = Type.GetTypeFromProgID("Excel.Application");
            return (type != null);
        }

        /// <summary>
        /// Create a new Excel file with a PivotTable connected to the server/database/cube provided
        /// </summary>
        /// <param name="serverName"></param>
        /// <param name="databaseName"></param>
        /// <param name="cubeName"></param>
        /// <param name="exceptionAction">Action that processes any exception - the function will return false, this is a way to manage logging/telemetry</param>
        /// <returns>true if the operation completes without errors, otherwise false (any exception is removed and the function returns false)</returns>
         public static bool CreateInstanceWithPivotTableViaInterop(string serverName, string databaseName, string cubeName, out string excelVersion, Action<Exception> exceptionAction )
         {
            var connectionString = ModelHelper.GetOleDbConnectionString(serverName, databaseName);
            var connectionName = $"AnalyzeInExcel [{ serverName }].[{ databaseName }].[{ cubeName }]";
            var commandText = cubeName;
            var pivotTableName = $"AnalyzeInExcelPivotTable";

            try
            {
                Excel.Application app = new Excel.Application();

                // Return Excel Version for diagnostics
                excelVersion = app.Version;
                
                // Create a new workbook
                var workbook = app.Workbooks.Add();

                // Create the connection
                var workbookConnection = workbook.Connections.Add(
                    Name: connectionName,
                    Description: "",
                    ConnectionString: $"OLEDB;{ connectionString }",
                    CommandText: commandText,
                    lCmdtype: 1
                    );

                // Create the pivotcache
                var pivotCache = workbook.PivotCaches().Create(
                    SourceType: Excel.XlPivotTableSourceType.xlExternal,
                    SourceData: workbookConnection
                    );
                pivotCache.RefreshOnFileOpen = false;

                // Get the active worksheet
                var worksheet = workbook.ActiveSheet;

                // Create the PivotTable
                var pivotTable = pivotCache.CreatePivotTable(
                    TableDestination: worksheet.Range["A1"],
                    TableName: pivotTableName,
                    ReadData: false
                    );

                // Show Excel
                app.Visible = true;

                // Set Excel window as foreground window
                var hwnd = app.Hwnd;
                SetForegroundWindow((IntPtr)hwnd);  // Note Hwnd is declared as int
            }
            catch (Exception ex)
            {
                // In case of error simply fails the request and forward the exception
                exceptionAction?.Invoke(ex);
                excelVersion = "<unknown>";
                return false;
            }
            return true;
        }
        public static bool CreateInstanceWithPivotTableViaDispatch(string serverName, string databaseName, string cubeName, out string excelVersion, Action<Exception> exceptionAction)
        {
            string connectionString = ModelHelper.GetOleDbConnectionString(serverName, databaseName);
            string connectionName = $"AnalyzeInExcel [{ serverName }].[{ databaseName }].[{ cubeName }]";
            string commandText = cubeName;
            string pivotTableName = $"AnalyzeInExcelPivotTable";
            string foundExcelVersion = null;

            try
            {
                var type = Type.GetTypeFromProgID("Excel.Application");
                dynamic app = Activator.CreateInstance(type);

                // Return Excel Version for diagnostics
                excelVersion = app.Version;
                foundExcelVersion = excelVersion;

                // Create a new workbook
                dynamic workbook = app.Workbooks.Add();

                // Create the connection
                dynamic workbookConnection = workbook.Connections.Add(
                    Name: connectionName,
                    Description: "",
                    ConnectionString: $"OLEDB;{ connectionString }",
                    CommandText: commandText,
                    lCmdtype: 1
                    );

                // Create the pivotcache
                dynamic pivotCache = workbook.PivotCaches().Create(
                    SourceType: Excel.XlPivotTableSourceType.xlExternal,
                    SourceData: workbookConnection
                    );
                pivotCache.RefreshOnFileOpen = false;

                // Get the active worksheet
                dynamic worksheet = workbook.ActiveSheet;

                // Create the PivotTable
                dynamic pivotTable = pivotCache.CreatePivotTable(
                    TableDestination: worksheet.Range["A1"],
                    TableName: pivotTableName,
                    ReadData: false
                    );

                // Show Excel
                app.Visible = true;

                // Set Excel window as foreground window
                long hwnd = app.Hwnd;
                SetForegroundWindow((IntPtr)hwnd);  // Note Hwnd is declared as int
            }
            catch (Exception ex)
            {
                // In case of error simply fails the request and forward the exception
                exceptionAction?.Invoke(ex);
                excelVersion = foundExcelVersion ?? "<unknown>";
                return false;
            }
            return true;
        }

    }
}
