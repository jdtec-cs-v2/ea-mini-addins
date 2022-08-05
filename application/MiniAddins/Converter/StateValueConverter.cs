using MiniAddins.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace MiniAddins.Converter
{
    /// <summary>
    /// 导出状态code-name的转换
    /// </summary>
    public class StateValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {

            string resKey;
            string resValue;

            Dictionary<ExportState, string> keys = new Dictionary<ExportState, string>();
            keys.Add(ExportState.NoExport, "multilang_290");
            keys.Add(ExportState.Already, "multilang_291");
            keys.Add(ExportState.HasChange, "multilang_292");

            resKey = keys[(ExportState)value];

            resValue = (string)(Application.Current.Properties["mainUserControl"] as UserControl).Resources[resKey];
            return resValue;
            
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => null;
    }
}
