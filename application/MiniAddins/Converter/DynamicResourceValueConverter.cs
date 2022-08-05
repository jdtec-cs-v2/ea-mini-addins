using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace MiniAddins.Converter
{
    public class DynamicResourceValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {

            string resKey;
            string resValue;

            resKey = $"{value}";
            resValue = (string)(Application.Current.Properties["mainUserControl"] as UserControl).Resources[resKey];

            return resValue;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => null;
    }
}
