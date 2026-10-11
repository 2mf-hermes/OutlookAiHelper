using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Adapters.Mail
{
    /// <summary>
    /// Maps a source kind to the reader that can serve it. Kept in one place so that
    /// adding the Graph reader is a change here and nothing in the UI moves.
    /// </summary>
    public static class MailReaderFactory
    {
        /// <summary>
        /// False for a kind this build can name but cannot read yet. The settings file is
        /// plain JSON a user can hand-edit, so a source can legitimately point at a kind
        /// with no reader; callers must fall back rather than throw.
        ///
        /// Both halves of the answer are required -- a registration the app can use and a
        /// reader that exists -- so what the UI is told can never disagree with what Create
        /// is able to hand back.
        /// </summary>
        public static bool IsImplemented(MailSourceKind kind, string graphClientId)
        {
            if (kind == MailSourceKind.LocalOutlook)
            {
                return true;
            }

            if (kind != MailSourceKind.Graph)
            {
                return false;
            }

            return GraphReaderAvailable && GraphApp.HasClientId(graphClientId);
        }

        /// <summary>
        /// The same question for callers that resolved no client id of their own, asked with
        /// the one this build shipped with.
        /// </summary>
        public static bool IsImplemented(MailSourceKind kind)
        {
            return IsImplemented(kind, GraphApp.DefaultClientId);
        }

        /// <summary>
        /// False until the Graph reader itself lands in this folder, and a switch rather than
        /// a scattered set of checks so that the day it lands is a one line change. Until
        /// then no caller can be offered a source that Create would refuse to build.
        /// </summary>
        private static readonly bool GraphReaderAvailable = false;

        /// <summary>
        /// Returns null when the kind has no reader yet. The local reader is passed in
        /// instead of being constructed here, so the app owns exactly one Outlook COM
        /// session no matter how often a reader is resolved.
        /// </summary>
        public static IMailReader Create(MailSourceKind kind, IMailReader local)
        {
            if (kind == MailSourceKind.LocalOutlook)
            {
                return local;
            }

            return null;
        }
    }
}
