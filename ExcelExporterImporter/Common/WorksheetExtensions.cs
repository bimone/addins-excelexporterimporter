using OfficeOpenXml;

namespace ExcelExporterImporter.Common
{
    internal static class WorksheetExtensions
    {
        public static void ApplyDefaultProtection(this ExcelWorksheet worksheet)
        {
            worksheet.Protection.IsProtected = true;

            worksheet.Protection.AllowSelectUnlockedCells = true;
            worksheet.Protection.AllowSelectLockedCells = true;

            worksheet.Protection.AllowFormatCells = true;
            worksheet.Protection.AllowFormatRows = true;
            worksheet.Protection.AllowFormatColumns = true;

            worksheet.Protection.AllowInsertColumns = false;
            worksheet.Protection.AllowInsertRows = false;
            worksheet.Protection.AllowInsertHyperlinks = false;

            worksheet.Protection.AllowDeleteColumns = false;
            worksheet.Protection.AllowDeleteRows = false;

            worksheet.Protection.AllowSort = true;
            worksheet.Protection.AllowAutoFilter = true;

            worksheet.Protection.AllowPivotTables = false;
            worksheet.Protection.AllowEditObject = false;
            worksheet.Protection.AllowEditScenarios = false;
        }
    }
}