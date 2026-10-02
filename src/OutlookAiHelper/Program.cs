using System;
using System.Windows;
using System.Windows.Threading;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper
{
    public static class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            AppDomain.CurrentDomain.UnhandledException += OnUnhandled;
            Adapters.Logging.FileLogger.Info("OutlookAiHelper start version=" + AppInfo.Version);
            try
            {
                var app = new System.Windows.Application();
                app.ShutdownMode = ShutdownMode.OnMainWindowClose;
                app.DispatcherUnhandledException += OnDispatcherUnhandled;
                var window = new UI.MainWindow();
                app.MainWindow = window;
                app.Run(window);
            }
            catch (Exception ex)
            {
                ShowFatal(ex);
            }
        }

        private static void OnDispatcherUnhandled(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            ShowFatal(e.Exception);
            e.Handled = true;
        }

        private static void OnUnhandled(object sender, UnhandledExceptionEventArgs e)
        {
            var ex = e.ExceptionObject as Exception;
            ShowFatal(ex ?? new Exception(Convert.ToString(e.ExceptionObject)));
        }

        private static void ShowFatal(Exception ex)
        {
            try
            {
                var text = ex == null ? "unknown" : ex.ToString();
                Adapters.Logging.FileLogger.Error("Fatal", ex);
                MessageBox.Show(
                    "OutlookAiHelper " + AppInfo.Version + " encountered a problem.\n\n"
                    + text,
                    "OutlookAiHelper",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            catch (Exception)
            {
            }
        }
    }
}
