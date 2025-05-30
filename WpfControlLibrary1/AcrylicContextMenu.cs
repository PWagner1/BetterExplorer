using System;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using BetterExplorerControls.Helpers;
using WPFUI.Win32;
using ContextMenu = System.Windows.Controls.ContextMenu;

namespace BetterExplorerControls {
  public class AcrylicContextMenu : ContextMenu {
    static AcrylicContextMenu() {
      DefaultStyleKeyProperty.OverrideMetadata(typeof(AcrylicContextMenu), new FrameworkPropertyMetadata(typeof(AcrylicContextMenu)));
    }

    public AcrylicContextMenu() { }

    protected override void OnOpened(RoutedEventArgs e) {
      base.OnOpened(e);
      var hwnd = (HwndSource)PresentationSource.FromVisual(this);
      hwnd.CompositionTarget.BackgroundColor = Color.FromArgb(0, 0, 0, 0);
      this.Background = Brushes.Transparent;
      AcrylicHelper.SetBlur(hwnd.Handle, AcrylicHelper.AccentFlagsType.Window, AcrylicHelper.AccentState.ACCENT_ENABLE_ACRYLICBLURBEHIND, (uint)Application.Current.Resources["SystemAcrylicTint"]);
      var nonClientArea = new Dwmapi.MARGINS(new Thickness(0));
      var preference = (Int32)Dwmapi.DWM_WINDOW_CORNER_PREFERENCE.DWMWCP_ROUND;
      Dwmapi.DwmSetWindowAttribute(hwnd.Handle, Dwmapi.DWMWINDOWATTRIBUTE.DWMWA_WINDOW_CORNER_PREFERENCE, ref preference, sizeof(uint));
      var pv = 2;
      Dwmapi.DwmSetWindowAttribute(hwnd.Handle, Dwmapi.DWMWINDOWATTRIBUTE.DWMWA_TRANSITIONS_FORCEDISABLED, ref pv, sizeof(uint));
      Dwmapi.DwmExtendFrameIntoClientArea(hwnd.Handle, ref nonClientArea);
    }

    private IntPtr WndProcHooked(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam, ref bool handled) {

      if (msg == 0x0021) {
        handled = true;
        return new IntPtr(0x0003);
      } else return IntPtr.Zero;
    }
  }
}
