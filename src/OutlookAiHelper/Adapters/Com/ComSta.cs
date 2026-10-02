using System;
using System.Threading;

namespace OutlookAiHelper.Adapters.Com
{
    /// <summary>
    /// Outlook Object Model is STA-affine. Calling it from Task.Run (MTA) can hard-kill the process.
    /// All COM work must run through this helper.
    /// </summary>
    public static class ComSta
    {
        public static T Run<T>(Func<T> func)
        {
            if (func == null)
            {
                throw new ArgumentNullException("func");
            }

            T result = default(T);
            Exception error = null;
            var thread = new Thread(() =>
            {
                try
                {
                    result = func();
                }
                catch (Exception ex)
                {
                    error = ex;
                }
            });
            thread.IsBackground = true;
            thread.Name = "OutlookCom-STA";
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (error != null)
            {
                throw error;
            }

            return result;
        }

        public static void Run(Action action)
        {
            Run<object>(() =>
            {
                action();
                return null;
            });
        }
    }
}
