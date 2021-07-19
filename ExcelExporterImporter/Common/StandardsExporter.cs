using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Autodesk.Revit.DB;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelExporterImporter.Common
{
    public partial class StandardsExporter
    {
        private const int StandardGuidColumn = 1;
        private const int StandardGuidRow = 1;
        private readonly CancellationToken cancellationToken;

        /// <summary>
        ///     Modifies the value of the cancellationToken variable
        /// </summary>
        /// <param name="cancellationToken"></param>
        public StandardsExporter(CancellationToken cancellationToken)
        {
            this.cancellationToken = cancellationToken;
        }

        /// <summary>
        ///     Export the table for style lines
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="workbook"></param>
        private void ExportLineStyles(Document doc, ExcelWorkbook workbook)
        {
            //We will search for the categories that contain the lines of styles
            var c = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);

            var worksheet = AddStyleSheet(workbook, Constants.StandardsLineStyles, false);

            ExportCategory(c, worksheet, 3, false);
            InsertStandardId(worksheet, Constants.StandardsGroupItemLineStylesUniqueId);
        }

        /// <summary>
        ///     Export the table for object styles
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="workbook"></param>
        private void ExportObjectStyles(Document doc, ExcelWorkbook workbook)
        {
            // We will search the categories for Annotation
            var categories = doc.Settings.Categories.Cast<Category>()
                .Where(c => c.GetCategoryType() == CategoryType.Annotation).OrderBy(c => c.Name);
            var worksheet = AddStyleSheet(workbook, Constants.StandardsAnnotationObjects, false);
            var row = 3;
            //We export for each category
            foreach (var category in categories)
            {
                row = ExportCategory(category, worksheet, row, false);
                if (cancellationToken.IsCancellationRequested) return;
            }

            InsertStandardId(worksheet, Constants.StandardsGroupItemAnnotationObjectsUniqueId);
            if (cancellationToken.IsCancellationRequested) return;
            //Addition of a new sheet for the Models Objects
            worksheet = AddStyleSheet(workbook, Constants.StandardsModelObjects, true);
            //We will search the categories for Model
            categories = doc.Settings.Categories.Cast<Category>().Where(c => c.GetCategoryType() == CategoryType.Model)
                .OrderBy(c => c.Name);
            //We export for each category
            row = 3;
            foreach (var category in categories)
            {
                row = ExportCategory(category, worksheet, row, true);
                if (cancellationToken.IsCancellationRequested) return;
            }

            InsertStandardId(worksheet, Constants.StandardsGroupItemModelObjectsUniqueId);
            if (cancellationToken.IsCancellationRequested) return;
            //Adding a new sheet for the Analytical model object
            worksheet = AddStyleSheet(workbook, Constants.StandardsAnalyticalModelObjects, false);
            //We will search the categories for Analytical Model
            categories = doc.Settings.Categories.Cast<Category>()
                .Where(c => c.GetCategoryType() == CategoryType.AnalyticalModel).OrderBy(c => c.Name);
            row = 3;
            //We export for each category
            foreach (var category in categories) row = ExportCategory(category, worksheet, row, false);
            InsertStandardId(worksheet, Constants.StandardsGroupItemAnalyticalModelObjectsUniqueId);
        }

        /// <summary>
        ///     Export categories
        /// </summary>
        /// <param name="category"></param>
        /// <param name="worksheet"></param>
        /// <param name="currentRow"></param>
        /// <param name="hasLineWeightCut"></param>
        /// <returns>return line count</returns>
        private int ExportCategory(Category category, ExcelWorksheet worksheet, int currentRow, bool hasLineWeightCut)
        {
            //We will search the list of subcategories
            var subcats = category.SubCategories;
            var row = currentRow;
            var colorColumn = hasLineWeightCut ? 5 : 4; //If hasLineWeightCut is True, the value is 5, otherwise it's 4
            //Fill in the line with category information
            worksheet.Cells[row, 1].Value = category.Id;
            worksheet.Cells[row, 2].Value = category.Name;
            //Enter the thickness of the line used for the Projection style
            worksheet.Cells[row, 3].Value = category.GetLineWeight(GraphicsStyleType.Projection);
            if (hasLineWeightCut)
                //Enter the thickness of the line used for the Cut style
                worksheet.Cells[row, 4].Value = category.GetLineWeight(GraphicsStyleType.Cut);
            //Valid if the line has a defined color
            if (category.LineColor.IsValid)
            {
                //We recover the different color values
                var color = string.Format("{0}, {1}, {2}", category.LineColor.Red, category.LineColor.Green,
                    category.LineColor.Blue);
                //We insert in a cell the value of the color
                worksheet.Cells[row, colorColumn].Value = color;
            }

            row++;
            //We retrieve the elements of the sub-category list
            foreach (var subCategory in subcats.Cast<Category>().OrderBy(c => c.Name))
            {
                if (cancellationToken.IsCancellationRequested) return currentRow;
                worksheet.Cells[row, 1].Value = subCategory.Id;
                worksheet.Cells[row, 2].Value = "      |----  " + subCategory.Name;
                worksheet.Cells[row, 3].Value = subCategory.GetLineWeight(GraphicsStyleType.Projection);

                if (hasLineWeightCut) worksheet.Cells[row, 4].Value = subCategory.GetLineWeight(GraphicsStyleType.Cut);
                if (subCategory.LineColor.IsValid)
                {
                    var color = string.Format("{0}, {1}, {2}", subCategory.LineColor.Red, subCategory.LineColor.Green,
                        subCategory.LineColor.Blue);
                    worksheet.Cells[row, colorColumn].Value = color;
                }

                row++;
            }

            //We automatically adjust the width of the columns
            RevitUtilities.AutoFitAllCol(worksheet);
            worksheet.Column(1).Hidden = true;
            return row;
        }

        /// <summary>
        ///     Export table styles of sheets
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="worksheetName"></param>
        /// <param name="addLineWeightCut"></param>
        /// <returns></returns>
        private ExcelWorksheet AddStyleSheet(ExcelWorkbook workbook, string worksheetName, bool addLineWeightCut)
        {
            var worksheet = workbook.Worksheets.Add(worksheetName);
            worksheet.ApplyDefaultProtection();

            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[2, 1].Value = Resources.ID;

            worksheet.Cells[1, 2].Value = "Category";
            worksheet.Cells[2, 2].Value = Resources.Category;

            worksheet.Cells[1, 3].Value = "Line Weight - Projection";
            worksheet.Cells[2, 3].Value = Resources.LineWeightProjection;

            worksheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Column(3).Style.Locked = false;
            worksheet.Column(4).Style.Locked = false;

            if (addLineWeightCut)
            {
                worksheet.Cells[1, 4].Value = "Line Weight - Cut";
                worksheet.Cells[2, 4].Value = Resources.LineWeightCut;

                worksheet.Cells[1, 5].Value = "Line Color (R,G,B)";
                worksheet.Cells[2, 5].Value = Resources.LineColorRGB;
                worksheet.Column(5).Style.Locked = false;
            }
            else
            {
                worksheet.Cells[1, 4].Value = "Line Color (R,G,B)";
                worksheet.Cells[2, 4].Value = Resources.LineColorRGB;
            }

            RevitUtilities.LockColumn(worksheet, 1);
            RevitUtilities.LockColumn(worksheet, 2);
            RevitUtilities.LockRow(worksheet, 1);
            RevitUtilities.FormattingTheHeader(worksheet, 2);
            worksheet.Column(1).Hidden = true;
            worksheet.Row(1).Hidden = true;
            return worksheet;
        }

        /// <summary>
        ///     Export the table with the family list
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="workbook"></param>
        private void ExportFamiliesList(Document doc, ExcelWorkbook workbook)
        {
            //****************************************This table is not importable*********************************************************
            //Creating a new sheet with the name Family Listing
            var worksheet = workbook.Worksheets.Add(Constants.StandardsFamilyListing);
            //Applies the default protection to the sheet
            worksheet.ApplyDefaultProtection();
            //Creating the table header
            worksheet.Cells[1, 1].Value = Resources.Name;
            //Locks all columns
            RevitUtilities.LockAllColumns(worksheet);
            //Apply formatting to header line
            RevitUtilities.FormattingTheHeader(worksheet, 1);
            //We are looking for categories of families
            var familiesCategories = new FilteredElementCollector(doc)
                .WherePasses(new ElementClassFilter(typeof(Family))).OrderBy(f => f.Name)
                .GroupBy(f => ((Family) f).FamilyCategory.Name).Select(f => new {CategoryName = f.Key, Families = f});
            //Initialization of the starting line
            var row = 2;
            foreach (var cat in familiesCategories.OrderBy(fc => fc.CategoryName))
            {
                if (cancellationToken.IsCancellationRequested) return;
                //Insert the name of the family categories
                worksheet.Cells[row, 1].Value = cat.CategoryName;
                foreach (Family family in cat.Families)
                {
                    //We insert the surnames
                    row++;
                    worksheet.Cells[row, 1].Value = "      |----  " + family.Name;
                }

                row++;
            }

            //We adjust the columns automatically
            RevitUtilities.AutoFitAllCol(worksheet);
            //Insert the Id of the sheet
            InsertStandardId(worksheet, Constants.StandardsGroupItemFamilyListingUniqueId);
            //Insertion of message indicating sheet cannot be imported
            RevitUtilities.InsertMsgNotBeImported(worksheet);
        }

        /// <summary>
        ///     Insert a line at the beginning of the sheet that will contain the group ID
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="standardsGroupItemUniqueId"></param>
        private void InsertStandardId(ExcelWorksheet worksheet, string standardsGroupItemUniqueId)
        {
            worksheet.InsertRow(StandardGuidRow, 1);
            worksheet.Cells[StandardGuidRow, StandardGuidColumn].Value = standardsGroupItemUniqueId;
            worksheet.Row(StandardGuidRow).Hidden = true;
        }

        /// <summary>
        ///     Export the table with project information
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="workbook"></param>
        /// <param name="parametersSettings"></param>
        private void ExportProjectInformation(Document doc, ExcelWorkbook workbook,
            ParametersSettings parametersSettings)
        {
            var parametersList = new Dictionary<string, string>();
            var appliedParameters = parametersSettings.ParametersTranslations
                .Where(p => p.Location == "*" || p.Location == Constants.StandardsProjectInformation).ToList();
            var projectInformation = doc.ProjectInformation;
            //Add a new sheet
            var worksheet = workbook.Worksheets.Add(Constants.StandardsProjectInformation);
            //Adding the default protection
            worksheet.ApplyDefaultProtection();
            //Creating the header
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[2, 1].Value = Resources.ID;

            worksheet.Cells[1, 2].Value = "Properties";
            worksheet.Cells[2, 2].Value = Resources.Properties;

            worksheet.Cells[1, 3].Value = "PValue";
            worksheet.Cells[2, 3].Value = Resources.Value;

            RevitUtilities.LockColumn(worksheet, 1);
            RevitUtilities.LockColumn(worksheet, 2);

            worksheet.Column(3).Style.Locked = false; //The third column is unlocked
            //The starting line number is indicated
            var row = 3;
            //Inserting rows with values
            foreach (var parameter in projectInformation.Parameters.Cast<Parameter>()
                .OrderBy(p => p.Definition.ParameterGroup).ThenBy(p => p.Definition.Name))
            {
                if (cancellationToken.IsCancellationRequested) return;

                if (!RevitUtilities.CanExportParameter(parameter, parametersSettings.IgnoredParameters,
                    "Project Information"))
                    continue;

                if (parametersList.Any(c => c.Value == parameter.Definition.Name) || parameter.IsReadOnly ||
                    parameter.Definition.ParameterType != ParameterType.Text)
                    continue;

                worksheet.Cells[row, 1].Value = parameter.Id;
                worksheet.Cells[row, 2].Value = parameter.Definition.Name;
#if REVIT2021
                    var format = RevitUtilities.GetUnitTypeSymbol(doc, parameter.Definition.GetSpecTypeId());
#else
                var format = RevitUtilities.GetUnitTypeSymbol(doc, parameter.Definition.UnitType);
#endif
                format = format.IndexOf(" ") > 0 ? format.Replace(" ", " \"") + "\"" : format;
                if (!string.IsNullOrEmpty(format)) worksheet.Cells[row, 3].Style.Numberformat.Format = format;
                if (!parameter.IsReadOnly)
                {
                    worksheet.Cells[row, 3].Style.Locked = false;
                }
                else
                {
                    worksheet.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Styles.BackgroundColor.CellLocked);
                    worksheet.Cells[row, 3].Style.Font.Color.SetColor(Styles.FontColor.CellLocked);
                }

                parametersList.Add(parameter.Definition.Name, parameter.Definition.Name);

                var parameterValue = RevitUtilities.GetParameterValue(doc, parameter);

                var appliedTranslations = appliedParameters.Where(p => p.Name == parameter.Definition.Name).ToList();
                worksheet.Cells[row, 3].Value = appliedTranslations.Any()
                    ? RevitUtilities.TranslateValueToText(Convert.ToString(parameterValue), appliedTranslations)
                    : parameterValue;

                row++;
            }

            //Formatting the header
            RevitUtilities.FormattingTheHeader(worksheet, 2);
            //We adjust the columns automatically
            RevitUtilities.AutoFitAllCol(worksheet);
            RevitUtilities.LockRow(worksheet, 1);
            worksheet.Row(1).Hidden = true;
            worksheet.Column(1).Hidden = true; //We mask the first column
            //Insert the Id of the sheet
            InsertStandardId(worksheet, Constants.StandardsGroupItemProjectInformationUniqueId);
        }

        /// <summary>
        ///     Export the table with the project parameters
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="workbook"></param>
        private void ExportProjectParametersDefinition(Document doc, ExcelWorkbook workbook)
        {
            //*****************************************************************************
            //*                      This table cannot be imported                        *
            //*****************************************************************************
            //Creating a new sheet with the name Project parameters
            var worksheet = workbook.Worksheets.Add(Constants.StandardsProjectParameters);
            //Application of default security
            worksheet.ApplyDefaultProtection();
            //Creating the table header
            worksheet.Cells[1, 1].Value = "GUID";
            worksheet.Cells[2, 1].Value = Resources.GUID;

            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Cells[2, 2].Value = Resources.Name;

            worksheet.Cells[1, 3].Value = "Group";
            worksheet.Cells[2, 3].Value = Resources.Group;

            worksheet.Cells[1, 4].Value = "Type";
            worksheet.Cells[2, 4].Value = Resources.Type;

            worksheet.Cells[1, 5].Value = "Instance Binding";
            worksheet.Cells[2, 5].Value = Resources.InstanceBinding;

            worksheet.Cells[1, 6].Value = "Visible";
            worksheet.Cells[2, 6].Value = Resources.Visible;

            worksheet.Cells[1, 7].Value = "Is Shared";
            worksheet.Cells[2, 7].Value = Resources.IsShared;

            worksheet.Cells[1, 8].Value = "Owner Group Name";
            worksheet.Cells[2, 8].Value = Resources.OwnerGroupName;

            worksheet.Cells[1, 9].Value = "Categories";
            worksheet.Cells[2, 9].Value = Resources.Categories;
            //Lock all columns
            RevitUtilities.LockAllColumns(worksheet);
            worksheet.Row(1).Hidden = true;
            RevitUtilities.FormattingTheHeader(worksheet, 2);
            //Initialization of the starting line
            var row = 3;
            //Retrieving values to add in the table
            var map = doc.ParameterBindings;
            var it = map.ForwardIterator();
            it.Reset();
#if REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021
                while (it.MoveNext())
                {
                    if (this.cancellationToken.IsCancellationRequested)
                    {
                        return ;
                    }
                    var eleBinding = it.Current as ElementBinding;
                    var insBinding = eleBinding as InstanceBinding;
                    var def = (InternalDefinition)it.Key;
                    if (def != null)
                    {
                        var sharedParameterElement = doc.GetElement(def.Id) as SharedParameterElement;
                        var shared = sharedParameterElement != null;
                        string sOwnerGroupName = string.Empty;
                        if(shared)
                        {
                            Element OwnerParameterElement = doc.GetElement(sharedParameterElement.OwnerViewId);
                            if(OwnerParameterElement != null)
                            {
                                Element eGroup = doc.GetElement(OwnerParameterElement.GroupId);
                                if(eGroup != null)
                                {
                                    if(eGroup.Name != null)
                                    {
                                        sOwnerGroupName = eGroup.Name;
                                    }
                                }
                            }
                        }
                        worksheet.Cells[row, 1].Value =
 shared ? sharedParameterElement.GuidValue.ToString() : string.Empty;
                        worksheet.Cells[row, 2].Value = def.Name;
                        worksheet.Cells[row, 3].Value = def.ParameterGroup;
                        worksheet.Cells[row, 4].Value = def.ParameterType;
                        worksheet.Cells[row, 5].Value = insBinding != null;
                        worksheet.Cells[row, 6].Value = def.Visible;
                        worksheet.Cells[row, 7].Value = shared;
                        worksheet.Cells[row, 8].Value = sOwnerGroupName;
                        worksheet.Cells[row, 9].Value =
 string.Join(",", eleBinding.Categories.Cast<Category>().Select(c => c.Name).ToArray());
                        row++;
                    }
                }
#endif
            RevitUtilities.AutoFitAllCol(worksheet);
            RevitUtilities.InsertMsgNotBeImported(worksheet);
            //Inserting the Id of the sheet
            InsertStandardId(worksheet, Constants.StandardsGroupItemProjectParametersSettingsUniqueId);
        }

        /// <summary>
        ///     Export the table with shared project parameters
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="workbook"></param>
        private void ExportSharedParametersDefinition(Document doc, ExcelWorkbook workbook)
        {
            //*****************************************************************************
            //*                      This table cannot be imported                        *
            //*****************************************************************************
            var map = doc.ParameterBindings;
            var it = map.ForwardIterator();
            var ListSharedParam = new List<InternalDefinition>();
            it.Reset();
            while (it.MoveNext())
            {
                if (cancellationToken.IsCancellationRequested) return;
                var eleBinding = it.Current as ElementBinding;
                var insBinding = eleBinding as InstanceBinding;
                var def = (InternalDefinition) it.Key;
                if (def != null)
                {
#if REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021
                        SharedParameterElement sharedParameterElement =
 doc.GetElement(def.Id) as SharedParameterElement;
                        if(sharedParameterElement != null)
                        {
                            ListSharedParam.Add(def);
                        }
#endif
                }
            }

            if (ListSharedParam.Count > 0)
            {
                var worksheet = workbook.Worksheets.Add(Constants.StandardsProjectSharedParameters);
                worksheet.ApplyDefaultProtection();
                //Creating the table header
                worksheet.Cells[1, 1].Value = "GUID";
                worksheet.Cells[2, 1].Value = Resources.GUID;

                worksheet.Cells[1, 2].Value = "Name";
                worksheet.Cells[2, 2].Value = Resources.Name;

                worksheet.Cells[1, 3].Value = "Group";
                worksheet.Cells[2, 3].Value = Resources.Group;

                worksheet.Cells[1, 4].Value = "Type";
                worksheet.Cells[2, 4].Value = Resources.Type;

                worksheet.Cells[1, 5].Value = "Visible";
                worksheet.Cells[2, 5].Value = Resources.Visible;

                //Lock all columns
                RevitUtilities.LockAllColumns(worksheet);
                worksheet.Row(1).Hidden = true;
                RevitUtilities.FormattingTheHeader(worksheet, 2);
                //Initialization of the departure line
                var row = 3;
#if REVIT2017 || REVIT2018 || REVIT2019 || REVIT2020 || REVIT2021
                    foreach (InternalDefinition iItem in ListSharedParam)
                    {
                        SharedParameterElement sItem = doc.GetElement(iItem.Id) as SharedParameterElement;
                        worksheet.Cells[row, 1].Value = sItem.GuidValue;
                        worksheet.Cells[row, 2].Value = sItem.Name;
                        worksheet.Cells[row, 3].Value = iItem.ParameterGroup;
                        worksheet.Cells[row, 4].Value = iItem.ParameterType.ToString();
                        worksheet.Cells[row, 5].Value = iItem.Visible;
                        row++;
                    }
#endif
                RevitUtilities.AutoFitAllCol(worksheet);
                RevitUtilities.InsertMsgNotBeImported(worksheet);
                //Inserting the Id of the sheet
                InsertStandardId(worksheet, Constants.StandardsGroupItemProjectSharedParametersSettingsUniqueId);
            }
        }

        /// <summary>
        ///     Exports the selected table by calling the correct method
        /// </summary>
        /// <param name="standardGroupId"></param>
        /// <param name="doc"></param>
        /// <param name="workbook"></param>
        /// <param name="parametersSettings"></param>
        public void ExportStandard(string standardGroupId, Document doc, ExcelWorkbook workbook,
            ParametersSettings parametersSettings)
        {
            switch (standardGroupId)
            {
                case Constants.StandardsGroupLineStylesUniqueId: //Line Styles
                    ExportLineStyles(doc, workbook);
                    break;

                case Constants.StandardsGroupObjectStylesUniqueId: //Model object
                    ExportObjectStyles(doc, workbook);
                    break;

                case Constants.StandardsGroupFamilyListingUniqueId: //Family Listing - Read Only
                    ExportFamiliesList(doc, workbook);
                    break;

                case Constants.StandardsGroupSharedParametersUniqueId: //Shared Parameters - Read Only
                    ExportSharedParametersDefinition(doc, workbook);
                    break;

                case Constants.StandardsGroupProjectParametersUniqueId: //Projet Parameters - Read Only
                    ExportProjectParametersDefinition(doc, workbook);
                    break;

                case Constants.StandardsGroupProjectInformationUniqueId: //Project Information
                    ExportProjectInformation(doc, workbook, parametersSettings);
                    break;
            }
        }
    }
}