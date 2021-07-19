using Autodesk.Revit.DB;

namespace ExcelExporterImporter.Common
{
    public static class ParameterUnitConverter
    {
        /// <summary>
        ///     Converts the number to the correct unit for export
        /// </summary>
        /// <param name="param"></param>
        /// <param name="scheduleField"></param>
        /// <returns></returns>
        public static double AsProjectUnitTypeDouble(this Parameter param, ScheduleField scheduleField = null)
        {
            var imperialValue = param.AsDouble();
            var document = param.Element.Document;
            double dResult = 0;
#if REVIT2021
                var fo = document.GetUnits().GetFormatOptions(param.Definition.GetSpecTypeId());
                //Condition for if the user did not use the default settings
                if (scheduleField != null)
                {
                    var foValue = scheduleField.GetFormatOptions();
                    if (!foValue.UseDefault)
                    {
                        fo = foValue;
                    }
                }
                dResult = UnitUtils.ConvertFromInternalUnits(imperialValue, fo.GetUnitTypeId());
#else
            var fo = document.GetUnits().GetFormatOptions(param.Definition.UnitType);
            //Condition for if the user did not use the default settings
            if (scheduleField != null)
            {
                var foValue = scheduleField.GetFormatOptions();
                if (!foValue.UseDefault) fo = foValue;
            }

            dResult = UnitUtils.ConvertFromInternalUnits(imperialValue, fo.DisplayUnits);
#endif
            return dResult;
        }

        /// <summary>
        ///     Converts the number to the correct unit for import
        /// </summary>
        /// <param name="param"></param>
        /// <param name="valueToConvert"></param>
        /// <param name="scheduleField"></param>
        /// <returns></returns>
        public static double ToProjectUnitType(this Parameter param, double valueToConvert,
            ScheduleField scheduleField = null)
        {
            var document = param.Element.Document;
            double dResult = 0;
#if REVIT2021
                var fo = document.GetUnits().GetFormatOptions(param.Definition.GetSpecTypeId());
                if (scheduleField != null)
                {
                    var foValue = scheduleField.GetFormatOptions();
                    if (!foValue.UseDefault)
                    {
                        fo = foValue;
                    }
                }   
                dResult = UnitUtils.ConvertToInternalUnits(valueToConvert, fo.GetUnitTypeId());
#else
            var fo = document.GetUnits().GetFormatOptions(param.Definition.UnitType);
            if (scheduleField != null)
            {
                var foValue = scheduleField.GetFormatOptions();
                if (!foValue.UseDefault) fo = foValue;
            }

            dResult = UnitUtils.ConvertToInternalUnits(valueToConvert, fo.DisplayUnits);
#endif
            return dResult;
        }
    }
}