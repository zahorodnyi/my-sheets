using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Layout;
using Avalonia.Media;

namespace MySheets.UI.Converters;

public class AlignmentToIconConverter : IValueConverter {
    private const string AlignLeftPath = "M3,21h18v-2H3V21z M3,17h12v-2H3V17z M3,13h18v-2H3V13z M3,9h12V7H3V9z M3,3v2h18V3H3z";
    private const string AlignCenterPath = "M7,21h10v-2H7V21z M3,17h18v-2H3V17z M7,13h10v-2H7V13z M3,9h18V7H3V9z M7,3v2h10V3H7z";
    private const string AlignRightPath = "M3,21h18v-2H3V21z M9,17h12v-2H9V17z M3,13h18v-2H3V13z M9,9h12V7H9V9z M3,3v2h18V3H3z";

    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture) {
        if (value is not HorizontalAlignment alignment) {
            return StreamGeometry.Parse(AlignLeftPath);
        }

        string pathData = alignment switch {
            HorizontalAlignment.Center => AlignCenterPath,
            HorizontalAlignment.Right => AlignRightPath,
            _ => AlignLeftPath 
        };

        return StreamGeometry.Parse(pathData);
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) {
        throw new NotImplementedException();
    }
}