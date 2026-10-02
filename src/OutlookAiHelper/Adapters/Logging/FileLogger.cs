using System;
using System.IO;

namespace OutlookAiHelper.Adapters.Logging
{
    public static class FileLogger
    {
        private static readonly object Gate = new object();

        public static string LogDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OutlookAiHelper",
                    "logs");
            }
        }

        public static void Info(string message)
        {
            Write("INFO", message, null);
        }

        public static void Error(string message, Exception ex)
        {
            Write("ERROR", message, ex);
        }

        private static void Write(string level, string message, Exception ex)
        {
            try
            {
                lock (Gate)
                {
                    Directory.CreateDirectory(LogDirectory);
                    var path = Path.Combine(LogDirectory, "app-" + DateTime.Now.ToString("yyyyMMdd") + ".log");
                    var line = DateTime.Now.ToString("o") + " [" + level + "] " + message;
                    if (ex != null)
                    {
                        line += " | " + ex.GetType().Name + ": " + ex.Message + " | " + ex.StackTrace;
                        if (ex.InnerException != null)
                        {
                            line += " | inner " + ex.InnerException.GetType().Name + ": " + ex.InnerException.Message;
                        }
                    }

                    File.AppendAllText(path, line + Environment.NewLine);
                }
            }
            catch (Exception)
            {
            }
        }

        public static string LatestLogPath()
        {
            try
            {
                var path = Path.Combine(LogDirectory, "app-" + DateTime.Now.ToString("yyyyMMdd") + ".log");
                return File.Exists(path) ? path : null;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
