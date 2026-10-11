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
        /// </summary>
        public static bool IsImplemented(MailSourceKind kind)
        {
            return kind == MailSourceKind.LocalOutlook;
        }

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
