using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ShellControls;
public class TabIndexConverter : IValueConverter
{
  public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
  {
    var tabItem = value as TabItem;
    var fromContainer = ItemsControl.ItemsControlFromItemContainer(tabItem)?.ItemContainerGenerator;

    var items = fromContainer?.Items.Cast<TabItem>().Where(x => x.Visibility == Visibility.Visible).ToList();
    var count = items?.Count() ?? 0;

    var index = items?.IndexOf(tabItem) ?? -2;
    if (index == -1 || index == -2) {
      return String.Empty;
    }

    if (index < count - 1) {
      var nextItem = items[index + 1];
      if (nextItem != null && (nextItem.IsSelected || nextItem.IsMouseOver)) {
        return "Collapse";
      }
    }

    return String.Empty;
  }

  public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
  {
    return DependencyProperty.UnsetValue;
  }
}
