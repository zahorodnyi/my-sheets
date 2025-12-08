using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace MySheets.UI.Converters;

public class StringToBrushConverter : IValueConverter {
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture) {
        if (value is string colorStr && !string.IsNullOrEmpty(colorStr)) {
            try {
                if (colorStr.Equals("Transparent", StringComparison.OrdinalIgnoreCase)) {
                    return Brushes.Transparent;
                }
                return Brush.Parse(colorStr);
            }
            catch {
                return Brushes.Transparent;
            }
        }
        return Brushes.Black;
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) {
        throw new NotImplementedException();
    }
}