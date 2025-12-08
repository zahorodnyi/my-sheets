using System;
using System.Globalization;
using Avalonia.Data.Converters;

namespace MySheets.UI.Converters;

public class ColorEqualityConverter : IValueConverter {
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture) {
        if (value is string color1 && parameter is string color2) {
            return string.Equals(color1, color2, StringComparison.OrdinalIgnoreCase);
        }
        return false;
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) {
        throw new NotImplementedException();
    }
}