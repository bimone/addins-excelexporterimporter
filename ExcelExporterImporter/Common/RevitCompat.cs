using Autodesk.Revit.DB;

namespace ExcelExporterImporter.Common
{
    internal static class RevitCompat
    {
        // ElementId.IntegerValue was replaced by ElementId.Value in Revit 2024.
        // Returning int keeps the original dictionary keys and enum casts working.
        public static int IntId(this ElementId id)
        {
#if REVIT2023 || REVIT2024
            return id.IntegerValue;
#else
            return (int)id.Value;
#endif
        }
    }
}
