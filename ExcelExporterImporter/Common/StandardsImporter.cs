using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Autodesk.Revit.DB;
using OfficeOpenXml;

namespace ExcelExporterImporter.Common
{
    public class StandardsImporter
    {
        private const int FieldHeaderRow = 2;
        private const int FirstDataRow = FieldHeaderRow + 2;
        private readonly CancellationToken cancellationToken;

        /// <summary>
        ///     Modifies the value of the cancellationToken variable
        /// </summary>
        /// <param name="cancellationToken"></param>
        public StandardsImporter(CancellationToken cancellationToken)
        {
            this.cancellationToken = cancellationToken;
        }

        /// <summary>
        ///     Call the method that imports the table that contains the line styles
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="worksheet"></param>
        /// <param name="progress"></param>
        public void ImportLineStyles(Document doc, ExcelWorksheet worksheet, Progress progress)
        {
            var c = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            ImportWorksheetAsCategories(doc, new List<Category> {c}, worksheet, progress);
        }

        /// <summary>
        ///     Call the method that imports the table that contains the annotation objects
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="worksheet"></param>
        /// <param name="progress"></param>
        public void ImportAnnotationObjects(Document doc, ExcelWorksheet worksheet, Progress progress)
        {
            var categories =
                doc.Settings.Categories.Cast<Category>()
                    .Where(c => c.GetCategoryType() == CategoryType.Annotation);
            ImportWorksheetAsCategories(doc, categories, worksheet, progress);
        }

        /// <summary>
        ///     Call the method that imports the table that contains the model objects
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="worksheet"></param>
        /// <param name="progress"></param>
        public void ImportModelObjects(Document doc, ExcelWorksheet worksheet, Progress progress)
        {
            var categories =
                doc.Settings.Categories.Cast<Category>()
                    .Where(c => c.GetCategoryType() == CategoryType.Model);

            ImportWorksheetAsCategories(doc, categories, worksheet, progress);
        }

        /// <summary>
        ///     Call the method that imports the table that contains analytical model objects
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="worksheet"></param>
        /// <param name="progress"></param>
        public void ImportAnalyticalModelObjects(Document doc, ExcelWorksheet worksheet, Progress progress)
        {
            var categories =
                doc.Settings.Categories.Cast<Category>()
                    .Where(c => c.GetCategoryType() == CategoryType.AnalyticalModel);
            ImportWorksheetAsCategories(doc, categories, worksheet, progress);
        }

        /// <summary>
        ///     Imports the table that contains project information
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="worksheet"></param>
        /// <param name="progress"></param>
        /// <param name="parametersSettings"></param>
        public void ImportProjectInformation(Document doc, ExcelWorksheet worksheet, Progress progress,
            ParametersSettings parametersSettings)
        {
            var projectInformation = doc.ProjectInformation;
            var cols = worksheet.Dimension.Columns;
            var rows = worksheet.Dimension.Rows;
            var changedSharedParameters = new Dictionary<Guid, Parameter>();
            var appliedParameters = parametersSettings.ParametersTranslations
                .Where(p => p.Location == "*" || p.Location == "Project Information").ToList();

            if (cols == 3 && rows >= FirstDataRow)
            {
                var parameterName = worksheet.Cells[FirstDataRow, 2].Value;
                var parameter = RevitUtilities.GetElementParameter(projectInformation, parameterName.ToString());
                if (parameter != null)
                {
                    var transaction = new Transaction(doc);
                    transaction.Start("Import from Excel");
                    try
                    {
                        for (var rowIndex = FirstDataRow; rowIndex <= rows; rowIndex++)
                        {
                            if (cancellationToken.IsCancellationRequested)
                            {
                                transaction.RollBack();
                                return;
                            }

                            if (string.IsNullOrEmpty(Convert.ToString(worksheet.Cells[rowIndex, 1].Value)))
                                continue;

                            parameterName = worksheet.Cells[rowIndex, 2].Value;
                            parameter = RevitUtilities.GetElementParameter(projectInformation,
                                parameterName.ToString());
                            if (parameter != null)
                            {
                                progress.SetStatus(string.Format("Importing parameter: {0} ({1})",
                                    parameter.Definition.Name, parameter.Id));

                                var cellVal = worksheet.Cells[rowIndex, 3].Value;

                                var appliedTranslations = appliedParameters
                                    .Where(p => p.Name == parameter.Definition.Name).ToList();
                                if (appliedTranslations.Any())
                                    cellVal = RevitUtilities.TranslateTextToValue(Convert.ToString(cellVal),
                                        appliedTranslations);
                                RevitUtilities.SetParameterValue(parameter, cellVal);
                            }

                            progress.Increment(1);
                        }

                        transaction.Commit();
                    }
                    catch (Exception)
                    {
                        transaction.RollBack();
                        throw;
                    }
                }
            }
        }

        /// <summary>
        ///     Gets the category
        /// </summary>
        /// <param name="id"></param>
        /// <param name="categories"></param>
        /// <returns>Category</returns>
        private static Category GetCategory(string id, IEnumerable<Category> categories)
        {
            foreach (var category in categories)
            {
                if (category.Id.IntegerValue == Convert.ToInt32(id))
                    return category;

                if (category.SubCategories != null)
                {
                    var cat = GetCategory(id, category.SubCategories.Cast<Category>());
                    if (cat != null)
                        return category;
                }
            }

            return null;
        }

        /// <summary>
        ///     Import table
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="categories"></param>
        /// <param name="worksheet"></param>
        /// <param name="progress"></param>
        private void ImportWorksheetAsCategories(Document doc, IEnumerable<Category> categories,
            ExcelWorksheet worksheet, Progress progress)
        {
            var cols = worksheet.Dimension.Columns;
            var rows = worksheet.Dimension.Rows;

            if (cols >= 2 && rows >= FirstDataRow)
            {
                var testId = worksheet.Cells[FirstDataRow, 1].Value;
                if (GetCategory(Convert.ToString(testId), categories) != null)
                {
                    var transaction = new Transaction(doc);
                    transaction.Start("Import from Excel");
                    try
                    {
                        for (var rowIndex = FirstDataRow; rowIndex <= rows; rowIndex++)
                        {
                            if (cancellationToken.IsCancellationRequested)
                            {
                                transaction.RollBack();
                                return;
                            }

                            if (string.IsNullOrEmpty(Convert.ToString(worksheet.Cells[rowIndex, 1].Value)))
                                continue;

                            var importCategory =
                                doc.Settings.Categories.get_Item(
                                    (BuiltInCategory) Convert.ToInt32(worksheet.Cells[rowIndex, 1].Value));
                            if (importCategory != null)
                            {
                                var projectionLineWeight = Convert.ToInt32(worksheet.Cells[rowIndex, 3].Value);
                                if (projectionLineWeight > 0)
                                    importCategory.SetLineWeight(projectionLineWeight, GraphicsStyleType.Projection);

                                var hasLineWeightCut = Convert.ToString(worksheet.Cells[FieldHeaderRow, 4].Value) ==
                                                       "Line Weight - Cut";
                                if (importCategory.IsCuttable && hasLineWeightCut)
                                {
                                    var cutLineWeight = Convert.ToInt32(worksheet.Cells[rowIndex, 4].Value);
                                    if (cutLineWeight > 0)
                                        importCategory.SetLineWeight(cutLineWeight, GraphicsStyleType.Cut);
                                }

                                var colorString =
                                    Convert.ToString(worksheet.Cells[rowIndex, hasLineWeightCut ? 5 : 4].Value);
                                var c = RevitUtilities.ConvertRgbToColor(colorString);
                                if (c != null) importCategory.LineColor = c;
                            }

                            progress.Increment(1);
                        }

                        transaction.Commit();
                    }
                    catch (Exception)
                    {
                        transaction.RollBack();
                        throw;
                    }
                }
            }
        }
    }
}