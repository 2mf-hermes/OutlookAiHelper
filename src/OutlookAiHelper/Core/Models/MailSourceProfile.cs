using System;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// Which backing service a mail source reads from.
    ///
    /// Stored as text in the settings file rather than as the enum's number: a file
    /// written by a newer build that knows a kind this build does not would otherwise
    /// deserialize to 0, which is LocalOutlook, and the app would quietly read the wrong
    /// mailbox. Text that fails to parse falls back to local instead.
    /// </summary>
    public enum MailSourceKind
    {
        /// <summary>The Outlook desktop client on this PC, through COM.</summary>
        LocalOutlook,

        /// <summary>A mailbox read from Microsoft Graph, with no Outlook running.</summary>
        Graph
    }

    /// <summary>
    /// One mailbox the app can read. The list of these lives in AppSettings and exactly
    /// one is active at a time; every scan reads only that one.
    ///
    /// There is no credential field here on purpose: a Graph source signs in through the
    /// device-code flow and the resulting refresh token is kept out of the settings file,
    /// which is plain JSON under %LocalAppData%.
    /// </summary>
    public sealed class MailSourceProfile
    {
        /// <summary>
        /// Fixed id of the implicit local source, so it stays the same across reinstalls
        /// and any stored reference to it keeps resolving.
        /// </summary>
        public const string LocalId = "local-outlook";

        public MailSourceProfile()
        {
            Id = Guid.NewGuid().ToString("N");
            Kind = MailSourceKind.LocalOutlook;
            Name = string.Empty;
            AccountUpn = string.Empty;
        }

        public string Id { get; set; }

        public MailSourceKind Kind { get; set; }

        /// <summary>
        /// What the user called this source. Empty is the normal state: the source list
        /// falls back to a translated label, so a source nobody renamed stays correct
        /// after the UI language changes.
        /// </summary>
        public string Name { get; set; }

        /// <summary>Signed-in address for a Graph source; empty for the local one.</summary>
        public string AccountUpn { get; set; }

        /// <summary>
        /// The local source every install has. A settings file that predates mail sources
        /// resolves to this, which is what keeps the upgrade invisible.
        /// </summary>
        public static MailSourceProfile Local()
        {
            return new MailSourceProfile
            {
                Id = LocalId,
                Kind = MailSourceKind.LocalOutlook,
                Name = string.Empty,
                AccountUpn = string.Empty
            };
        }

        /// <summary>
        /// Outlook's object model is STA-affine; Graph is a plain HTTPS call. Asked here
        /// rather than at the call site so the UI never has to know what kind it holds.
        /// </summary>
        public bool RequiresSta
        {
            get { return Kind == MailSourceKind.LocalOutlook; }
        }

        public static string KindToText(MailSourceKind kind)
        {
            return kind == MailSourceKind.Graph ? "Graph" : "LocalOutlook";
        }

        public static MailSourceKind ParseKind(string text)
        {
            if (string.Equals(text, "Graph", StringComparison.OrdinalIgnoreCase))
            {
                return MailSourceKind.Graph;
            }

            return MailSourceKind.LocalOutlook;
        }
    }
}
