using System;
using System.Collections.Generic;
using System.Threading;
using Autodesk.Revit.DB;
using OfficeOpenXml;

namespace ExcelExporterImporter.Common
{
    public class ScheduleImporter
    {
        private const int FieldHeaderRow = 2;
        private const int FirstDataRow = FieldHeaderRow + 3;
        private readonly CancellationToken cancellationToken;

        /// <summary>
        ///     Modifies the value of the cancellationToken variable
        /// </summary>
        /// <param name="cancellationToken"></param>
        public ScheduleImporter(CancellationToken cancellationToken)
        {
            this.cancellationToken = cancellationToken;
        }

        /// <summary>
        ///     Starts the import process
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="worksheet"></param>
        /// <param name="schedule"></param>
        /// <param name="progress"></param>
        /// <param name="parametersSettings"></param>
        public void ImportViewSchedule(Document doc, ExcelWorksheet worksheet, ViewSchedule schedule, Progress progress,
            ParametersSettings parametersSettings)
        {
            var iColStart = 3;
            var cols = worksheet.Dimension.Columns; //Indicates the number of columns the table has
            var rows = worksheet.Dimension.Rows; //Indicates the number of rows the table has
            var fields = new Dictionary<int, ScheduleField>();
            var fieldsCount = schedule.Definition.GetFieldCount();
            var progressUnit = 1000 / rows;
            //================Create a list with parameters that are read-only=========================
            var readonlyParameters = RevitUtilities.GetListReadOnlyParamater(parametersSettings);
            if (cols >= iColStart && rows >= FirstDataRow)
            {
                var fieldDictionary = new Dictionary<string, ScheduleField>();
                //We will search the list of column names via the information from the schedules
                for (var fieldIndex = 0; fieldIndex < fieldsCount; fieldIndex++)
                {
                    var field = schedule.Definition.GetField(fieldIndex);
                    if (field.HasSchedulableField)
                    {
                        var fieldName = field.GetSchedulableField().GetName(doc);
                        var i = 1;
                        while (fieldDictionary.ContainsKey(fieldName))
                        {
                            fieldName = $"{field.GetName()}({i})";
                            i++;
                        }

                        fieldDictionary.Add(fieldName, field);
                    }
                }

                //We will search the list of column names in excel
                for (var columnIndex = iColStart; columnIndex <= cols; columnIndex++)
                {
                    var columnName = Convert.ToString(worksheet.Cells[FieldHeaderRow, columnIndex].Value);
                    //Validate the column name is valid with the names of nomanclature
                    if (fieldDictionary.ContainsKey(columnName)) fields.Add(columnIndex, fieldDictionary[columnName]);
                }

                var transaction = new Transaction(doc);
                transaction.Start("Schedule Import from Excel");
                try
                {
                    for (var rowIndex = FirstDataRow; rowIndex <= rows; rowIndex++)
                    {
                        //If we cancel, we must go back
                        if (cancellationToken.IsCancellationRequested)
                        {
                            transaction.RollBack();
                            return;
                        }

                        //We validate that the line has an id
                        if (string.IsNullOrEmpty(Convert.ToString(worksheet.Cells[rowIndex, 1].Value))) continue;
                        //We validate that the id is good
                        var element = doc.GetElement((string) worksheet.Cells[rowIndex, 1].Value);
                        if (element != null)
                        {
                            progress.SetStatus(
                                string.Format(Resources.ImportingElement, element.Name, element.UniqueId));
                            //We recover the phase
                            var pElement = RevitUtilities.GetElementPhase(doc, element);
                            foreach (var scheduleField in fields)
                            {
                                var cellVal = worksheet.Cells[rowIndex, scheduleField.Key].Value;
                                var parameterId = scheduleField.Value.ParameterId;

                                //We call the method that will get the parameters associated with the cell
                                var parameter =
                                    RevitUtilities.GetParameter(doc, element, scheduleField.Value, pElement);
                                //We call the method that will get the rights associated with the cell
                                var readonlyParameter = RevitUtilities.GetIsReadOnly(parameter, scheduleField.Value,
                                    readonlyParameters);
                                if (!readonlyParameter)
                                    RevitUtilities.SetParameterValue(parameter, cellVal, scheduleField.Value);
                            }
                        }

                        progress.Increment(progressUnit);
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