using Autodesk.Revit.DB;

namespace ExcelExporterImporter.Common
{
    public static class ScheduleExtensions
    {
        public static bool IsDisplayTypeTotals(this ScheduleField field)
        {
            return field.DisplayType == ScheduleFieldDisplayType.Totals;
        }
    }
}