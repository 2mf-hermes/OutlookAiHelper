using System;
using System.Net;

namespace OutlookAiHelper.Adapters.Update
{
    /// <summary>
    /// One place that opts the whole process into modern TLS, for both the release check
    /// and the asset download.
    ///
    /// This exists because of a real, reproduced failure: the default security protocol of
    /// a fresh .NET Framework process here is Ssl3|Tls, and api.github.com as well as the
    /// asset host refuse that handshake. The check used to set it for itself only, which
    /// left the download working *by accident* whenever a check had run first in the same
    /// process - and failing with a misleading "network error" when it had not.
    /// </summary>
    internal static class UpdateHttp
    {
        private const int Tls12 = 3072;

        private static readonly object Gate = new object();

        private static bool _configured;

        /// <summary>Idempotent; safe to call before every request.</summary>
        internal static void EnsureModernTls()
        {
            if (_configured)
            {
                return;
            }

            lock (Gate)
            {
                if (_configured)
                {
                    return;
                }

                try
                {
                    ServicePointManager.SecurityProtocol =
                        ServicePointManager.SecurityProtocol | (SecurityProtocolType)Tls12;
                }
                catch (Exception)
                {
                    // A runtime that does not know the Tls12 value: keep the platform default.
                }

                _configured = true;
            }
        }
    }
}
