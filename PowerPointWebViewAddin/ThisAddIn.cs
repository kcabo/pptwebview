using System;

namespace PowerPointWebViewAddin
{
    public partial class ThisAddIn
    {
        private OverlayWindow overlayWindow;

        private void ThisAddIn_Startup(
            object sender,
            EventArgs e
        )
        {
            overlayWindow =
                new OverlayWindow(
                    this.Application
                );

            overlayWindow.Show();
        }

        private void ThisAddIn_Shutdown(
            object sender,
            EventArgs e
        )
        {
            overlayWindow?.Close();
        }

        #region VSTO で生成されたコード

        private void InternalStartup()
        {
            this.Startup +=
                new EventHandler(
                    ThisAddIn_Startup
                );

            this.Shutdown +=
                new EventHandler(
                    ThisAddIn_Shutdown
                );
        }

        #endregion
    }
}