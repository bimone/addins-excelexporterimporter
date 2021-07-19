using System;
using System.Globalization;
using System.Windows.Data;

namespace ExcelExporterImporter.Converters
{
    [ValueConversion(typeof(Enum), typeof(bool))]
    public class EnumToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null) return false;
            var enumValue = value.ToString();
            var targetValue = parameter.ToString();
            var outputValue = enumValue.Equals(targetValue, StringComparison.InvariantCultureIgnoreCase);
            return outputValue;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null) return null;
            var useValue = (bool) value;
            var targetValue = parameter.ToString();
            if (useValue) return Enum.Parse(targetType, targetValue);
            return null;
        }
    }
}