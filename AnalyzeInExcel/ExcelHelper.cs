using System;

namespace AnalyzeInExcel
{
    public class ExcelHelper
    {
        /// <summary>
        /// Check whether Excel is installed
        /// </summary>
        /// <returns>true if Excel is installed</returns>
        public static bool IsExcelAvailable()
        {
            var type = Type.GetTypeFromProgID("Excel.Application");
            return (type != null);
        }
    }
}
