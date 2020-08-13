using System;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using System.Runtime.InteropServices;

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

         public static bool CreateInstanceWithPivotTable(string serverName, string databaseName, string cubeName, Action<Exception> exceptionAction )
        {
            const int XlLayoutRowType_xlCompactRow = 0; // Excel.XlLayoutRowType.xlCompactRow
            const int XlPivotTableSourceType_xlExternal = 2; // Excel.XlPivotTableSourceType.xlExternal
            const int XlPivotFieldRepeatLabels_xlRepeatLabels = 2; // Excel.XlPivotFieldRepeatLabels.xlRepeatLabels
            
            var connectionString = ModelHelper.GetOleDbConnectionString(serverName, databaseName);
            var connectionName = $"AnalyzeInExcel [{ serverName }].[{ databaseName }].[{ cubeName }]";
            var commandText = cubeName;
            var pivotTableName = $"AnalyzeInExcelPivotTable";

            var type = Type.GetTypeFromProgID("Excel.Application");
            if (type == null)
                return false;

            dynamic app = Activator.CreateInstance(type);
            try
            {
                try
                {
                    var workbook = app.Workbooks.Add();

                    var workbookConnection = workbook.Connections.Add(
                        Name: connectionName,
                        Description: "",
                        ConnectionString: $"OLEDB;{ connectionString }",
                        CommandText: commandText,
                        lCmdtype: 1
                        );

                    var pivotCache = workbook.PivotCaches().Create(
                        SourceType: XlPivotTableSourceType_xlExternal,
                        SourceData: workbookConnection
                        );

                    #region Configure PivotCache

                    pivotCache.RefreshOnFileOpen = false;

                    #endregion

                    var worksheet = workbook.ActiveSheet;

                    var pivotTable = pivotCache.CreatePivotTable(
                        TableDestination: worksheet.Range["A1"],
                        TableName: pivotTableName,
                        ReadData: false
                        );

                    #region Configure PivotTable

                    pivotTable.ColumnGrand = true;
                    pivotTable.HasAutoFormat = true;
                    pivotTable.DisplayErrorString = true;
                    pivotTable.DisplayNullString = true;
                    pivotTable.EnableDrilldown = true;
                    pivotTable.ErrorString = "";
                    pivotTable.MergeLabels = false;
                    pivotTable.NullString = "";
                    pivotTable.PageFieldOrder = 2;
                    pivotTable.PageFieldWrapCount = 0;
                    pivotTable.PreserveFormatting = true;
                    pivotTable.RowGrand = true;
                    pivotTable.PrintTitles = false;
                    pivotTable.RepeatItemsOnEachPrintedPage = true;
                    pivotTable.TotalsAnnotation = true;
                    pivotTable.CompactRowIndent = 1;
                    pivotTable.VisualTotals = false;
                    pivotTable.InGridDropZones = false;
                    pivotTable.DisplayFieldCaptions = true;
                    pivotTable.DisplayMemberPropertyTooltips = true;
                    pivotTable.DisplayContextTooltips = true;
                    pivotTable.ShowDrillIndicators = true;
                    pivotTable.PrintDrillIndicators = false;
                    pivotTable.DisplayEmptyRow = false;
                    pivotTable.DisplayEmptyColumn = false;
                    pivotTable.AllowMultipleFilters = false;
                    pivotTable.SortUsingCustomLists = true;
                    pivotTable.DisplayImmediateItems = true;
                    pivotTable.ViewCalculatedMembers = true;
                    pivotTable.EnableWriteback = false;
                    pivotTable.ShowValuesRow = false;
                    pivotTable.CalculatedMembersInFilters = true;
                    pivotTable.RowAxisLayout(XlLayoutRowType_xlCompactRow);
                    pivotTable.RepeatAllLabels(XlPivotFieldRepeatLabels_xlRepeatLabels);

                    #endregion

                    // Show Excel
                    app.Visible = true;

                    // Set Excel window as foreground window
                    var hwnd = app.Hwnd;
                    SetForegroundWindow((IntPtr)hwnd);  // Note Hwnd is declared as int
                }
                finally
                {
                    Marshal.ReleaseComObject(app);
                }
            }
            catch (Exception ex)
            {
                // In case of error simply fails the request and forward the exception
                exceptionAction?.Invoke(ex);
                return false;
            }
            return true;
        }
    }
}
