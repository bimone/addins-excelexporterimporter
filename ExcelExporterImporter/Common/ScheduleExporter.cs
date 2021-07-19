using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using Autodesk.Revit.DB;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelExporterImporter.Common
{
    public class ScheduleExporter
    {
        private const int ScheduleGuidColumn = 1;
        private const int ScheduleGuidRow = 1;
        private readonly CancellationToken cancellationToken;

        /// <summary>
        ///     Modifies the content of the cancellationToken variable
        /// </summary>
        /// <param name="cancellationToken">CancellationToken</param>
        /// <returns></returns>
        public ScheduleExporter(CancellationToken cancellationToken)
        {
            this.cancellationToken = cancellationToken;
        }

        /// <summary>
        ///     Function which allows to export the nomenclature as it is in revit
        /// </summary>
        /// <param name="schedule">Schedule information</param>
        /// <param name="worksheet">Excel file</param>
        /// <returns></returns>
        public void ExportViewScheduleBasic(ViewSchedule schedule, ExcelWorksheet worksheet)
        {
            var dt = new DataTable();
            //Definition of columns
            var fieldsCount = schedule.Definition.GetFieldCount();
            for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
            {
                var field = schedule.Definition.GetField(fieldIndex);
                if (field.IsHidden) continue;
                var fieldType = typeof(string);
                var columnName = field.ColumnHeading;
                var i = 1;
                while (dt.Columns.Contains(columnName))
                {
                    columnName = $"{field.GetName()}({i})";
                    i++;
                }

                dt.Columns.Add(columnName, fieldType);
            }

            //Content display
            var viewSchedule = schedule;
            var table = viewSchedule.GetTableData();
            var section = table.GetSectionData(SectionType.Body);
            var nRows = section.NumberOfRows;
            var nColumns = section.NumberOfColumns;
            if (nRows > 1)
                //Starts at 1 so as not to display the header
                for (var i = 1; i < nRows; i++)
                {
                    var data = dt.NewRow();
                    for (var j = 0; j < nColumns; j++)
                    {
                        object val = viewSchedule.GetCellText(SectionType.Body, i, j);
                        if (val.ToString() != "") data[j] = val;
                    }

                    dt.Rows.Add(data);
                }

            if (dt.Rows.Count > 0)
            {
                worksheet.Cells.LoadFromDataTable(dt, true);
                RevitUtilities.AutoFitAllCol(worksheet);
            }
        }

        /// <summary>
        ///     Export a schedule to an Excel file
        /// </summary>
        /// <param name="doc">Document</param>
        /// <param name="schedule">Schedule information</param>
        /// <param name="worksheet">Excel file</param>
        /// <param name="parametersSettings">ParametersSettings</param>
        /// <returns></returns>
        public void ExportViewSchedule(Document doc, ViewSchedule schedule, ExcelWorksheet worksheet,
            ParametersSettings parametersSettings)
        {
            var appliedParameters = parametersSettings.ParametersTranslations
                .Where(p => p.Location == "*" || p.Location == "ViewSchedule_" + schedule.Name).ToList();
            var LstParameter = new Dictionary<int, Parameter>();
            var LstLastTypeFamilly = new Dictionary<string, int>();
            var revitLinksElements = new FilteredElementCollector(doc, schedule.Id)
                .OfCategory(BuiltInCategory.OST_RvtLinks).ToElementIds();
            //---Gets the list of items---
            var collector = new FilteredElementCollector(doc, schedule.Id).WhereElementIsNotElementType();
            var iStartRow = 2; //Position of content insertion

            var bAnalyticalNodesShedule = false;
            var bRvtLinksShedule = false;
            if (BuiltInCategory.OST_AnalyticalNodes == (BuiltInCategory) schedule.Definition.CategoryId.IntegerValue)
                bAnalyticalNodesShedule = true;
            else if (BuiltInCategory.OST_RvtLinks == (BuiltInCategory) schedule.Definition.CategoryId.IntegerValue)
                bRvtLinksShedule = true;
            //Excluded revit links
            if (revitLinksElements.Any() && bRvtLinksShedule == false)
                collector = collector.Excluding(revitLinksElements);
            //----------------------------------     
            var fieldsList = new List<ScheduleField>();
            var dt = new DataTable();
            //=========================================Creating the excel table header=================================
            var fieldsCount = schedule.Definition.GetFieldCount();
            dt.Columns.Add("ID");
            dt.Columns.Add("FamilyAndType");
            for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
            {
                var field = schedule.Definition.GetField(fieldIndex);
                if (!field.HasSchedulableField)
                {
                    //continue;
                }

                if (!RevitUtilities.CanExportParameter(field, parametersSettings.IgnoredParameters,
                    "ViewSchedule_" + schedule.Name)) continue;
                var fieldType = typeof(string);
                if (field.CanTotal()) fieldType = typeof(double);
                fieldsList.Add(field);
                var columnName = field.GetName();

                var i = 1;
                while (dt.Columns.Contains(columnName))
                {
                    columnName = $"{field.GetName()}({i})";
                    i++;
                }

                dt.Columns.Add(columnName, fieldType);
            }

            //================================End of the creation of the header of the table excel======================
            //=================================Create a list with parameters that are read-only=========================
            var readonlyParameters = RevitUtilities.GetListReadOnlyParamater(parametersSettings);
            //=============================================Get each element line========================================
            var iRow = iStartRow;
            foreach (var element in collector)
            {
                if (element.Name == string.Empty && bAnalyticalNodesShedule == false) //Remove empty lines
                    continue;
                var col = 2;
                var data = dt.NewRow();
                data["ID"] = element.UniqueId;
                //=========================We will look for the type and family and note its position in the dictionary==========================
                //For ElementsType, only the values written on the last line of a type and family member used for the update. We then note the last line of each family type and we will lock the other cells to avoid errors
                var parameter_temp = element.get_Parameter((BuiltInParameter) (-1002052));
                if (parameter_temp != null)
                {
                    var elementType = doc.GetElement(parameter_temp.AsElementId()) as ElementType;
                    if (elementType != null)
                    {
                        var familyName = RevitUtilities.GetElementFamilyName(doc, elementType);
                        var sTypeNameFamilly = familyName.Trim() + ": " + elementType.Name.Trim();
                        data["FamilyAndType"] = sTypeNameFamilly;
                    }
                }

                //===========================================Get each of the values for the fields===========================
                var pElement = RevitUtilities.GetElementPhase(doc, element);
                foreach (var scheduleField in fieldsList) //Use the list of columns generated above
                {
                    if (cancellationToken.IsCancellationRequested) return;
                    //We call the method that will get the parameters associated with the cell
                    var parameter = RevitUtilities.GetParameter(doc, element, scheduleField, pElement);
                    if (parameter != null)
                        if (!LstParameter.ContainsKey(scheduleField.ParameterId.IntegerValue))
                            LstParameter.Add(scheduleField.ParameterId.IntegerValue, parameter);
                    //We call the method that will get the rights associated with the cell
                    var readonlyParameter = RevitUtilities.GetIsReadOnly(parameter, scheduleField, readonlyParameters);
                    //We call the method that will get the value associated with the cell
                    var cellVal =
                        RevitUtilities.GetParameterValue(parameter, scheduleField, doc, element, appliedParameters);
                    //Assign values and parameters to the cell
                    dt.Columns[col].ReadOnly = readonlyParameter;
                    data[col] = cellVal ?? DBNull.Value;
                    col++;
                }

                iRow += 1;
                //===================End of the recovery each of the values for the fields===========================
                dt.Rows.Add(data);
            }

            //========================================End of the Recovery of each element line==========================
            //=========================================Creation of sort and filter string===================================
            var sStringSort = RevitUtilities.GetStringSort(schedule, doc, fieldsList);
            var sStringFilter = RevitUtilities.GetStringFilter(schedule, doc);
            var sScheduleName = schedule.Name;
            //===========================================Application of filter==========================================
            if (!string.IsNullOrEmpty(sStringFilter))
            {
                dt.DefaultView.RowFilter = sStringFilter;
                dt = dt.DefaultView.ToTable(sScheduleName);
            }

            //===========================================Application of sort and filter=========================================
            if (!string.IsNullOrEmpty(sStringSort))
            {
                dt.DefaultView.Sort = sStringSort;
                dt = dt.DefaultView.ToTable(sScheduleName);
                var dtTemp = NaturalSorting.DataTableSort(dt, sStringSort, out var sMsgError).Copy();
                if (string.IsNullOrEmpty(sMsgError)) dt = dtTemp.Copy();
            }

            //======================================We get the positioning of the elements which corresponds to the last ElementType=========================
            iRow = iStartRow;
            foreach (DataRow vRow in dt.Rows)
            {
                var sFamilyType = vRow[1].ToString().Trim();
                if (sFamilyType != "")
                {
                    if (LstLastTypeFamilly.ContainsKey(sFamilyType))
                        LstLastTypeFamilly[sFamilyType] = iRow;
                    else
                        LstLastTypeFamilly.Add(sFamilyType, iRow);
                }

                iRow += 1;
            }

            //===========================================================================================================================================================================
            worksheet.Cells.LoadFromDataTable(dt, true);
            //Insert one row at the top to store schedule unique id
            worksheet.InsertRow(ScheduleGuidRow, 1);
            worksheet.Cells[ScheduleGuidRow, ScheduleGuidColumn].Value = schedule.UniqueId;
            //Hide the first column
            worksheet.Column(ScheduleGuidColumn).Hidden = true;
            //Hide the first row
            worksheet.Row(ScheduleGuidRow).Hidden = true;
            worksheet.View.FreezePanes(3, 1); //We freeze the menu
            FormatWorksheet(doc, schedule, worksheet, fieldsList, dt, LstParameter, LstLastTypeFamilly);
        }

        /// <summary>
        ///     This method changes the format of the table.
        /// </summary>
        /// <param name="doc">Document</param>
        /// <param name="schedule">Schedule information</param>
        /// <param name="worksheet">Excel file</param>
        /// <param name="fieldsList">List of columns</param>
        /// <param name="dt">Contents of the table</param>
        /// <param name="LstParameter">Parameter list</param>
        /// <param name="LstLastTypeFamilly">List of unlocked cells for item types</param>
        /// <returns></returns>
        private void FormatWorksheet(Document doc, ViewSchedule schedule, ExcelWorksheet worksheet,
            List<ScheduleField> fieldsList, DataTable dt, Dictionary<int, Parameter> LstParameter,
            Dictionary<string, int> LstLastTypeFamilly)
        {
            var iStartCol = 3;
            var iRowAjust = 1;
            var iStartRow = 3;
            var iTotalRows = worksheet.Dimension.Rows; //Gives the total number of lines
            var ListColHidden = new List<int> {1, 2}; //List which contains the numbers which must be hidden
            var ListColFormula = new List<int>(); //List of columns which contains a formula

            RevitUtilities.FormatingTable(worksheet);
            foreach (var scheduleField in fieldsList)
            {
                var fieldId = scheduleField.GetName();
                var colIndex = dt.Columns[fieldId].Ordinal + 1;
                //Hide the column
                if (scheduleField.IsHidden) ListColHidden.Add(colIndex);
                if (scheduleField.FieldType == ScheduleFieldType.Formula) ListColFormula.Add(colIndex);
                //Get the field format
                var format = "";
                var formatOptions = scheduleField.GetFormatOptions();
#if REVIT2021
                    if (!formatOptions.UseDefault && !formatOptions.GetSymbolTypeId().Empty())
                    {
                        var formatValueOptions = new FormatValueOptions();
                        formatValueOptions.SetFormatOptions(formatOptions);
                        format =
 UnitFormatUtils.Format(doc.GetUnits(), scheduleField.GetSpecTypeId(), 0, true, formatValueOptions);
                    }
                    else if (formatOptions.UseDefault && !scheduleField.GetSpecTypeId().Empty())
                    {
                        format = RevitUtilities.GetUnitTypeSymbol(doc, scheduleField.GetSpecTypeId());
                    }
#else
                if (!formatOptions.UseDefault && formatOptions.UnitSymbol != UnitSymbolType.UST_NONE)
                {
                    var formatValueOptions = new FormatValueOptions();
                    formatValueOptions.SetFormatOptions(formatOptions);
                    format = UnitFormatUtils.Format(doc.GetUnits(), scheduleField.UnitType, 0, true, false,
                        formatValueOptions);
                }
                else if (formatOptions.UseDefault && scheduleField.UnitType != UnitType.UT_Undefined)
                {
                    format = RevitUtilities.GetUnitTypeSymbol(doc, scheduleField.UnitType);
                }
#endif
                format = format.IndexOf(" ") > 0 ? format.Replace(" ", " \"") + "\"" : format;
                if (!string.IsNullOrEmpty(format)) worksheet.Column(colIndex).Style.Numberformat.Format = format;
                //Change the color of the column if it can not be modified
                if (dt.Columns[fieldId].ReadOnly || scheduleField.FieldType == ScheduleFieldType.Count ||
                    scheduleField.FieldType == ScheduleFieldType.Formula ||
                    scheduleField.FieldType == ScheduleFieldType.MaterialQuantity)
                {
                    RevitUtilities.LockColumn(worksheet, colIndex);
                }
                else
                {
                    worksheet.Column(colIndex).Style.Locked = false;
                    //===================We indicate the cells for the ElementType columns which can be modified and we lock the others==============
                    if (scheduleField.FieldType == ScheduleFieldType.ElementType && LstLastTypeFamilly.Count > 0)
                    {
                        RevitUtilities.FormattingColElementType(worksheet, colIndex);
                        for (var iRow = iStartRow; iRow <= iTotalRows; iRow++)
                            if (worksheet.Cells[iRow, 2].Value != null)
                            {
                                var sTypeAndFamilly = worksheet.Cells[iRow, 2].Value.ToString();
                                var iPosition = LstLastTypeFamilly[sTypeAndFamilly];
                                if (iPosition > 0)
                                {
                                    if (iPosition == iRow - iRowAjust)
                                    {
                                        //Unlock the cell
                                        RevitUtilities.UnlockCell(worksheet, iRow, colIndex);
                                        //Addition "" to not have 0 for null cell
                                        if (worksheet.Cells[iRow, colIndex].Value == null)
                                            worksheet.Cells[iRow, colIndex].Value = "";
                                    }
                                    else
                                    {
                                        //Adding a formula
                                        worksheet.Cells[iRow, colIndex].Formula =
                                            "=" + worksheet.Cells[iPosition + iRowAjust, colIndex].Address;
                                    }
                                }
                            }
                    }
                }
            }

            //We group the group lines
            var level = 0;
            var iLevelMax = schedule.Definition.GetSortGroupFields().Count;
            var dGroupFieldId = new Dictionary<string, int>();
            foreach (var scheduleSortGroupField in schedule.Definition.GetSortGroupFields())
            {
                level++;
                // We will get the name of the column
                var fieldId = schedule.Definition.GetField(scheduleSortGroupField.FieldId).GetSchedulableField()
                    .GetName(doc); //Give the Id of the column
                var column = dt.Columns[fieldId]; //It will look for the properties of the column
                if (column == null) continue;
                dGroupFieldId.Add(fieldId, level);
                var colIndex =
                    column.Ordinal + 1; //Indicates the column number. We do plus 1 in the first column this is the id
                var dic = new Dictionary<string, int>();
                var groupFirstRow = 3; //To start playback after declaring the title bar
                var bSubRow = false;
                for (var rowIndex = groupFirstRow; rowIndex <= iTotalRows; rowIndex++)
                {
                    var sIdRowIndex = worksheet.Cells[rowIndex, 1].Value + "";
                    if (sIdRowIndex == "")
                    {
                        bSubRow = true;
                        continue;
                    }

                    var sCellValue = "StringEmpty";
                    if (worksheet.Cells[rowIndex, colIndex].Value != null)
                    {
                        sCellValue = worksheet.Cells[rowIndex, colIndex].Value.ToString();
                    }
                    else if (worksheet.Cells[rowIndex, colIndex].Formula != null)
                    {
                        var sFormula = worksheet.Cells[rowIndex, colIndex].Formula;
                        if (sFormula != "")
                        {
                            var sCell = sFormula.Substring(1, sFormula.Length - 1);
                            sCellValue = worksheet.Cells[sCell].Value.ToString();
                        }
                    }

                    var sIdKey = sCellValue + "-level=" + level;
                    var bContainsKey = dic.ContainsKey(sIdKey);
                    if (!bContainsKey || bSubRow) //We check if the column name has already been inserted
                    {
                        bSubRow = false;
                        if (iTotalRows != rowIndex && !bContainsKey) //We validate that we haven't come to the end
                            dic.Add(sIdKey, level); //We add the value to the dictionary with its level

                        if (cancellationToken.IsCancellationRequested
                        ) //If the user has clicked on the cancel button we stop everything.
                            return;

                        if (dic.Count > 1 && (scheduleSortGroupField.ShowFooter || !schedule.Definition.IsItemized))
                        {
                            AddFooter(worksheet, rowIndex, colIndex, groupFirstRow, fieldsList, dt, level, iLevelMax,
                                dGroupFieldId);
                            rowIndex++;
                            iTotalRows++;
                        }

                        groupFirstRow = rowIndex;
                    }

                    if (iTotalRows == rowIndex && (scheduleSortGroupField.ShowFooter || !schedule.Definition.IsItemized)
                    ) //When we reach the last line, we create a foot
                        AddFooter(worksheet, rowIndex + 1, colIndex, groupFirstRow, fieldsList, dt, level, iLevelMax,
                            dGroupFieldId);
                    worksheet.Row(rowIndex).OutlineLevel =
                        level; //defines the current hierarchical level of the specified row or column.
                    if (!schedule.Definition.IsItemized) worksheet.Row(rowIndex).Collapsed = true;
                }
            }

            //===============================Creating the total line and formatting the header==============================
            if (level > 0)
            {
                var iFirstRow = 3;
                var iLastRow = worksheet.Dimension.Rows;
                var iNewRow = iLastRow + 1;

                worksheet.InsertRow(iNewRow, 1);
                for (var index = 0; index < fieldsList.Count(); index++)
                {
                    var field = fieldsList[index];
                    var fieldId = field.GetName();
                    var columnIndex = dt.Columns[fieldId].Ordinal + 1;
                    if (field.IsDisplayTypeTotals())
                        worksheet.Cells[iNewRow, columnIndex].Formula = string.Format("=SUBTOTAL(9,{0})",
                            worksheet.Cells[iFirstRow, columnIndex, iLastRow, columnIndex].Address);
                    if (field.FieldType == ScheduleFieldType.Count)
                        worksheet.Cells[iNewRow, columnIndex].Formula = string.Format("=SUBTOTAL(9,{0})",
                            worksheet.Cells[iFirstRow, columnIndex, iLastRow, columnIndex].Address);
                }

                //Format for the total line
                RevitUtilities.FormatingTotalRow(worksheet, iNewRow);
            }

            //Format for the header line
            var iHeaderRow = 2;
            //Adding the custom header and the type of value that the column must contain
            worksheet.InsertRow(3, 2);
            var iCol = iStartCol;
            foreach (var FieldItem in fieldsList)
            {
                worksheet.Cells[iHeaderRow + 1, iCol].Value = FieldItem.ColumnHeading;
                var sParamType = string.Empty;
                if (LstParameter.ContainsKey(FieldItem.ParameterId.IntegerValue))
                {
                    var vParam = LstParameter[FieldItem.ParameterId.IntegerValue];
                    sParamType = LstParameter[FieldItem.ParameterId.IntegerValue].Definition.ParameterType.ToString();
                    if (sParamType == "Invalid")
                        sParamType = vParam.StorageType.ToString();
                    else if (sParamType == "YesNo") sParamType = "TrueFalse";
                }

                worksheet.Cells[iHeaderRow + 2, iCol].Value = sParamType;
                iCol++;
            }

            worksheet.Row(iHeaderRow).Hidden = true;
            RevitUtilities.Unlock(worksheet, iHeaderRow + 1);
            if (ListColFormula.Count >= 1
            ) //If it has formula columns, change the color of the column and display which indicates that the formulas are not exportable.
            {
                foreach (var iColu in ListColFormula)
                    RevitUtilities.FormattingColFormula(worksheet, iStartRow, worksheet.Dimension.Rows, iColu);
                var iRowInsert = worksheet.Dimension.Rows + 2;
                var iColInsertEnd = worksheet.Dimension.Columns;
                worksheet.Cells[iRowInsert, 2].Value = Resources.WarningFormula;
                worksheet.Cells[iRowInsert, 2, iRowInsert, iColInsertEnd].Merge = true;
                worksheet.Cells[iRowInsert, 2, iRowInsert, iColInsertEnd].Style.WrapText = true;
                worksheet.Cells[iRowInsert, 2, iRowInsert, iColInsertEnd].Style.VerticalAlignment =
                    ExcelVerticalAlignment.Center;
                worksheet.Cells[iRowInsert, 2, iRowInsert, iColInsertEnd].Style.HorizontalAlignment =
                    ExcelHorizontalAlignment.Center;
                RevitUtilities.FormattingRowWarningFormula(worksheet, iRowInsert, 2, iRowInsert, iColInsertEnd);
                worksheet.Row(iRowInsert).Height *= 2;
            }

            RevitUtilities.FormattingTheHeader(worksheet, iHeaderRow + 1);
            RevitUtilities.LockRow(worksheet, iHeaderRow + 2);
            RevitUtilities.FormattingTheHeaderTypeField(worksheet, iHeaderRow + 2);
            RevitUtilities.LockRow(worksheet, iHeaderRow);
            RevitUtilities.FreezeRow(worksheet, iHeaderRow + 3);
            //=====================================End of the creation of the total line and formatting of the header====================================================
            //Finalize the Excel table
            worksheet.OutLineApplyStyle = true;
            worksheet.ApplyDefaultProtection();
            RevitUtilities.AutoFitAllCol(worksheet);
            //Hide the add columns in the ListColHidden list
            foreach (var iItemHidden in ListColHidden) worksheet.Column(iItemHidden).Hidden = true;
        }

        /// <summary>
        ///     Creation of the footer
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="groupColumnIndex"></param>
        /// <param name="groupFirstRow"></param>
        /// <param name="fieldsList"></param>
        /// <param name="dt"></param>
        /// <param name="iLevel"></param>
        /// <param name="iMaxLevel"></param>
        /// <param name="dGroupFieldId"></param>
        private static void AddFooter(ExcelWorksheet worksheet, int rowIndex, int groupColumnIndex, int groupFirstRow,
            List<ScheduleField> fieldsList, DataTable dt, int iLevel, int iMaxLevel,
            Dictionary<string, int> dGroupFieldId)
        {
            //Validation to insert a subtotal over the previous sub-total
            if (worksheet.Cells[rowIndex - 1, 1].Value + "" == "")
            {
                var iSubRow = iLevel - 1;
                rowIndex = rowIndex - iSubRow;
            }

            worksheet.InsertRow(rowIndex, 1);
            for (var index = 0; index < fieldsList.Count(); index++)
            {
                var field = fieldsList[index];
                var fieldId = field.GetName();
                var columnIndex = dt.Columns[fieldId].Ordinal + 1;
                var bFieldCount = false;
                if (field.IsDisplayTypeTotals())
                {
                    worksheet.Cells[rowIndex, columnIndex].Formula = string.Format("=SUBTOTAL(9,{0})",
                        worksheet.Cells[groupFirstRow, columnIndex, rowIndex - 1, columnIndex].Address);
                    bFieldCount = true;
                }

                if (field.FieldType == ScheduleFieldType.Count)
                {
                    worksheet.Cells[rowIndex, columnIndex].Formula = string.Format("=SUBTOTAL(9,{0})",
                        worksheet.Cells[groupFirstRow, columnIndex, rowIndex - 1, columnIndex].Address);
                    bFieldCount = true;
                }

                if (bFieldCount == false)
                    if (dGroupFieldId.ContainsKey(fieldId))
                        worksheet.Cells[rowIndex, columnIndex].Value = Convert
                            .ToString(worksheet.Cells[rowIndex - 1, columnIndex].Value + "").Replace("\"", "''");
            }

            if (iLevel == 1)
                RevitUtilities.FormatingLevel1(worksheet, rowIndex);
            else if (iLevel == 2) RevitUtilities.FormatingLevel2(worksheet, rowIndex);
        }
    }
}