using System;

namespace OutlookAiHelper.Adapters.Mail
{
    /// <summary>
    /// The fixed parts of the Microsoft Graph path: the app registration this public client
    /// identifies itself with, the endpoints, the one permission it asks for, and the exact
    /// message fields it reads.
    ///
    /// A public client has no secret, so a client id is public information and is never
    /// masked. The only value on this path that has to be protected is the refresh token that
    /// comes back from signing in: it is encrypted for the current user and never written to
    /// the settings file or the log.
    ///
    /// Nothing here talks to the network, so every statement below is pinned by ordinary unit
    /// tests rather than by a live call.
    /// </summary>
    public static class GraphApp
    {
        /// <summary>
        /// The registration the app ships with, so a cloud mailbox can work out of the box.
        ///
        /// Empty means "no registration yet": the cloud source stays unimplemented and the UI
        /// keeps offering only the local mailbox, instead of offering a sign-in that cannot
        /// succeed. Filling in the client id (a GUID) is the whole change needed to turn the
        /// cloud path on; the UI asks through MailReaderFactory.IsImplemented.
        /// </summary>
        public const string DefaultClientId = "";

        /// <summary>
        /// A registration that serves company tenants and personal Microsoft accounts is
        /// addressed as "common". A single-tenant registration could not be addressed this way,
        /// which is why the registration has to be multi-tenant plus personal.
        /// </summary>
        public const string Tenant = "common";

        private const string AuthorizeHost = "https://login.microsoftonline.com";
        private const string GraphHost = "https://graph.microsoft.com";

        public const string DeviceCodeUrl = AuthorizeHost + "/" + Tenant + "/oauth2/v2.0/devicecode";

        public const string TokenUrl = AuthorizeHost + "/" + Tenant + "/oauth2/v2.0/token";

        public const string MailFoldersUrl = GraphHost + "/v1.0/me/mailFolders";

        /// <summary>
        /// Read-only, and only message metadata: Mail.ReadBasic excludes the body, the body
        /// preview, the unique body, attachments and extended properties. The app never reads a
        /// message body and never writes to a mailbox, so nothing wider is needed. A test
        /// refuses any scope that is not on the read-only allowlist, so widening this cannot
        /// happen by accident.
        /// </summary>
        public static readonly string[] Scopes = { "openid", "profile", "offline_access", "Mail.ReadBasic" };

        /// <summary>
        /// The exact fields the app asks Graph for, kept next to the code that reads them so a
        /// field asked for and a field read cannot drift apart. A test refuses any entry that
        /// carries a message body or an attachment.
        /// </summary>
        public static readonly string[] SelectFields =
        {
            "id", "subject", "from", "receivedDateTime", "isRead", "flag", "importance", "meetingMessageType"
        };

        public static string ScopeText
        {
            get { return string.Join(" ", Scopes); }
        }

        public static string SelectText
        {
            get { return string.Join(",", SelectFields); }
        }

        /// <summary>
        /// True when the value can actually be used as a client id. A build whose value is empty
        /// or a leftover placeholder must not claim it can sign in, so the shape is checked
        /// rather than the string merely being non-empty.
        /// </summary>
        public static bool HasClientId(string clientId)
        {
            if (string.IsNullOrEmpty(clientId))
            {
                return false;
            }

            Guid parsed;
            return Guid.TryParse(clientId.Trim(), out parsed);
        }

        /// <summary>
        /// Where the messages of one folder live. The well-known name "inbox" is a valid folder
        /// id, so the cheap new-mail probe and a full scan share this one url shape.
        /// </summary>
        public static string MessagesUrl(string folderId)
        {
            if (string.IsNullOrEmpty(folderId))
            {
                return GraphHost + "/v1.0/me/messages";
            }

            return GraphHost + "/v1.0/me/mailFolders/" + Uri.EscapeDataString(folderId) + "/messages";
        }
    }
}
