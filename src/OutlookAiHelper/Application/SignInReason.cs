using OutlookAiHelper.Adapters.Mail;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.Application
{
    /// <summary>
    /// Turns the word a sign-in ended on into a sentence the settings page can show, and says
    /// whether a word is one this app has something better to say than the word itself.
    ///
    /// Kept beside the sign-in rather than in the window for the same reason
    /// MailSourceAvailability is: a rule a test can reach without a window. Every key below is
    /// spelled out as a literal so that the string audit can see it -- a key assembled from the
    /// word would compile and run, and a translation missing from one language would go unnoticed
    /// by the one check that exists to notice exactly that.
    ///
    /// A word this app does not know is shown as itself. "unauthorized_client" or "http-403" tells
    /// whoever is asked to look into it more than one soothing sentence would.
    /// </summary>
    internal static class SignInReason
    {
        /// <summary>
        /// What to show for one reason word. The two words that mean the browser step is still
        /// running read as waiting rather than as failures, so a row can put Reason on show
        /// whatever the last outcome was, without having to know the outcome as well.
        /// </summary>
        internal static string TextFor(string word)
        {
            if (word == GraphAuth.Pending || word == GraphAuth.SlowDown)
            {
                return Strings.T("source.signIn.waiting");
            }

            switch (word)
            {
                case GraphSignIn.NoClientIdReason:
                    return Strings.T("source.reason.no-client-id");
                case GraphSignIn.UnreadableReason:
                    return Strings.T("source.reason.unreadable");
                case GraphSignIn.OfflineReason:
                    return Strings.T("source.reason.offline");
                case GraphSignIn.NoCodeReason:
                    return Strings.T("source.reason.no-code");
                case GraphSignIn.NoRefreshTokenReason:
                    return Strings.T("source.reason.no-refresh-token");
                case ExpiredToken:
                    return Strings.T("source.reason.expired_token");
                case AccessDenied:
                    return Strings.T("source.reason.access_denied");
                default:
                    return word;
            }
        }

        /// <summary>
        /// True when there is a sentence for this word, so that a caller can tell "the app has
        /// nothing to add" from "the app has never heard of this".
        /// </summary>
        internal static bool IsKnown(string word)
        {
            return !string.IsNullOrEmpty(word) && TextFor(word) != word;
        }

        /// <summary>
        /// The service's own word for a code that outlived its lifetime. A literal here because the
        /// sign-in keeps the service's words as they arrived, so there is no constant of ours to
        /// point at; the test pins this string to the word the sign-in reports.
        /// </summary>
        private const string ExpiredToken = "expired_token";

        /// <summary>The service's word for a consent the user or their organisation refused.</summary>
        private const string AccessDenied = "access_denied";
    }
}
