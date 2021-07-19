using Autodesk.Revit.DB;

namespace ExcelExporterImporter.Common
{
    public static class CategoryExtensions
    {
        /// <summary>
        ///     Gets the type of category
        /// </summary>
        /// <param name="cat"></param>
        /// <returns>Category TYpe</returns>
        public static CategoryType GetCategoryType(this Category cat)
        {
            return cat.CategoryType;
        }
    }
}