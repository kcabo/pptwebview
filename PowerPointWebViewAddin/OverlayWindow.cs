using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointWebViewAddin
{
    public class OverlayWindow : Window
    {
        private readonly WebView2CompositionControl webView;
        private readonly PowerPoint.Application powerPoint;

        private const int GWL_EXSTYLE = -20;
        private const long WS_EX_TOOLWINDOW = 0x00000080;
        private const long WS_EX_APPWINDOW = 0x00040000;

        [DllImport(
            "user32.dll",
            EntryPoint = "GetWindowLongPtr",
            SetLastError = true
        )]
        private static extern IntPtr GetWindowLongPtr64(
            IntPtr hWnd,
            int nIndex
        );

        [DllImport(
            "user32.dll",
            EntryPoint = "GetWindowLong",
            SetLastError = true
        )]
        private static extern int GetWindowLong32(
            IntPtr hWnd,
            int nIndex
        );

        [DllImport(
            "user32.dll",
            EntryPoint = "SetWindowLongPtr",
            SetLastError = true
        )]
        private static extern IntPtr SetWindowLongPtr64(
            IntPtr hWnd,
            int nIndex,
            IntPtr dwNewLong
        );

        [DllImport(
            "user32.dll",
            EntryPoint = "SetWindowLong",
            SetLastError = true
        )]
        private static extern int SetWindowLong32(
            IntPtr hWnd,
            int nIndex,
            int dwNewLong
        );

        private static IntPtr GetWindowLongPtr(
            IntPtr hWnd,
            int nIndex
        )
        {
            if (IntPtr.Size == 8)
            {
                return GetWindowLongPtr64(
                    hWnd,
                    nIndex
                );
            }

            return new IntPtr(
                GetWindowLong32(
                    hWnd,
                    nIndex
                )
            );
        }

        private static void SetWindowLongPtr(
            IntPtr hWnd,
            int nIndex,
            IntPtr value
        )
        {
            if (IntPtr.Size == 8)
            {
                SetWindowLongPtr64(
                    hWnd,
                    nIndex,
                    value
                );
            }
            else
            {
                SetWindowLong32(
                    hWnd,
                    nIndex,
                    value.ToInt32()
                );
            }
        }

        public OverlayWindow(
            PowerPoint.Application powerPoint
        )
        {
            this.powerPoint = powerPoint;

            // Web側から後でサイズ変更される。
            Width = 200;
            Height = 60;

            // 最初は画面外
            Left = -10000;
            Top = -10000;

            WindowStyle =
                WindowStyle.None;

            ResizeMode =
                ResizeMode.NoResize;

            AllowsTransparency =
                true;

            Background =
                Brushes.Transparent;

            // Chrome等より常に前にはしない
            Topmost = false;

            // タスクバーに出さない
            ShowInTaskbar = false;

            // Show時にPowerPointから
            // フォーカスを奪わない
            ShowActivated = false;

            webView =
                new WebView2CompositionControl
                {
                    DefaultBackgroundColor =
                        System.Drawing.Color.Transparent
                };

            Content = webView;

            SourceInitialized +=
                ConfigureNativeWindow;

            Loaded +=
                InitializeWebView;
        }

        private void ConfigureNativeWindow(
            object sender,
            EventArgs e
        )
        {
            var helper =
                new WindowInteropHelper(this);

            IntPtr overlayHwnd =
                helper.Handle;

            /*
             * PowerPointのOwned Windowにする。
             *
             * PowerPointより前、
             * Chrome等より常時前ではない。
             */
            IntPtr powerPointHwnd =
                new IntPtr(
                    powerPoint.HWND
                );

            if (powerPointHwnd != IntPtr.Zero)
            {
                helper.Owner =
                    powerPointHwnd;
            }

            /*
             * ToolWindowとして扱い、
             * タスクバー/Alt+Tabへ出さない。
             */
            long style =
                GetWindowLongPtr(
                    overlayHwnd,
                    GWL_EXSTYLE
                )
                .ToInt64();

            style &=
                ~WS_EX_APPWINDOW;

            style |=
                WS_EX_TOOLWINDOW;

            SetWindowLongPtr(
                overlayHwnd,
                GWL_EXSTYLE,
                new IntPtr(style)
            );
        }

        private async void InitializeWebView(
            object sender,
            RoutedEventArgs e
        )
        {
            string userDataFolder =
                Environment.GetEnvironmentVariable(
                    "PPTADDIN_WEBVIEW2_DATA"
                );

            if (string.IsNullOrWhiteSpace(userDataFolder))
            {
                userDataFolder =
                    Path.Combine(
                        Environment.GetFolderPath(
                            Environment.SpecialFolder.LocalApplicationData
                        ),
                        "PowerPointWebViewAddin",
                        "WebView2"
                    );
            }

            Directory.CreateDirectory(
                userDataFolder
            );

            var environment =
                await CoreWebView2Environment
                    .CreateAsync(
                        null,
                        userDataFolder
                    );

            await webView
                .EnsureCoreWebView2Async(
                    environment
                );

            /*
             * PowerPoint COMそのものを
             * Web側へ公開。
             */
            webView.CoreWebView2
                .AddHostObjectToScript(
                    "ppt",
                    powerPoint
                );

            /*
             * Window操作だけを
             * 別Bridgeとして公開。
             */
            webView.CoreWebView2
                .AddHostObjectToScript(
                    "overlay",
                    new WindowBridge(this)
                );

            webView.Source =
                new Uri(
                    "http://127.0.0.1:5173"
                );
        }
    }
}
