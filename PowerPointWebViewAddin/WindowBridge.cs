using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;

namespace PowerPointWebViewAddin
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class WindowBridge
    {
        private readonly Window window;

        public WindowBridge(Window window)
        {
            this.window = window;
        }

        public void SetBounds(
            double screenX,
            double screenY,
            double width,
            double height
        )
        {
            RunOnUiThread(() =>
            {
                DpiScale dpi =
                    VisualTreeHelper.GetDpi(window);

                // PowerPointのscreen pixel
                // → WPFのDIPへ変換
                window.Left =
                    screenX / dpi.DpiScaleX;

                window.Top =
                    screenY / dpi.DpiScaleY;

                window.Width =
                    Math.Max(1, width);

                window.Height =
                    Math.Max(1, height);
            });
        }

        public void MoveOffscreen()
        {
            RunOnUiThread(() =>
            {
                // Window自体は生かしておく。
                // これによりWebView側のsetIntervalも止まらない。
                window.Left = -10000;
                window.Top = -10000;
            });
        }

        private void RunOnUiThread(Action action)
        {
            if (window.Dispatcher.CheckAccess())
            {
                action();
            }
            else
            {
                window.Dispatcher.Invoke(action);
            }
        }
    }
}