using System;
using System.Runtime.InteropServices;
using OutlookAiHelper.Adapters.Logging;

namespace OutlookAiHelper.Adapters.Outlook
{
    public static class OutlookAvailability
    {
        /// <summary>
        /// Attach to a RUNNING Outlook only. Never Activator.CreateInstance — that can launch Outlook
        /// without user intent and surprise-crash when profile/permissions are not ready.
        /// </summary>
        public static bool TryGetRunningApplication(out object application, out string reasonKey)
        {
            application = null;
            reasonKey = "error.outlook.not_detected";

            try
            {
                application = Marshal.GetActiveObject("Outlook.Application");
                if (application != null)
                {
                    reasonKey = null;
                    FileLogger.Info("Outlook attached via ROT");
                    return true;
                }
            }
            catch (COMException ex)
            {
                FileLogger.Error("Outlook not in ROT (not running or not ready)", ex);
                application = null;
                return false;
            }
            catch (Exception ex)
            {
                FileLogger.Error("Outlook GetActiveObject failed", ex);
                application = null;
                return false;
            }

            return false;
        }

        public static void Release(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch (Exception)
            {
            }
        }
    }
}
