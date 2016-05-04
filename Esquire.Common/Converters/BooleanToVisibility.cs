using System.Windows;
using System.Windows.Data;

namespace Esquire.Common.Converters
{
    public class BooleanToVisibility : IValueConverter
    {

        public object Convert(object value, System.Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            bool condition = (bool)value;
            return condition ? Visibility.Visible : Visibility.Collapsed;
        }

        //TODO: implement?
        public object ConvertBack(object value, System.Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new System.NotImplementedException();
        }
    }
}
