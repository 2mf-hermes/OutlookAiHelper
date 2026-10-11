using System;
using System.Collections.Generic;
using System.Globalization;
using OutlookAiHelper.Adapters.Logging;
using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Adapters.Mail
{
    /// <summary>
    /// Reads one cloud mailbox over Microsoft Graph, using the sign-in the device code flow stored
    /// for it.
    ///
    /// What this buys over the desktop reader is two things, and only two: a mailbox can be read on
    /// a machine whose Outlook is not signed in to it, and it can be read on a machine where Outlook
    /// is not running at all. Everything else -- which mail is urgent, what the list looks like --
    /// is decided above this class, so a mail read here and the same mail read through Outlook are
    /// indistinguishable to the rest of the app.
    ///
    /// Three rules shape the code:
    ///  - a failure is never reported as an empty mailbox. "Could not be read" and "nothing in range"
    ///    look the same to a caller that only counts rows, so a failed read stops the scan and says
    ///    why, through the same reason keys the desktop reader uses;
    ///  - nothing here writes. The permission this app asks for (Mail.ReadBasic) does not allow it,
    ///    the app asks only for metadata, and neither does this code;
    ///  - no token reaches a message, a log line or the settings file. Tokens live in the encrypted
    ///    store and in this object, and a failed sign-in is logged by its status and error code only.
    ///
    /// Not routed through the STA/COM marshalling the desktop reader needs: there is no COM object on
    /// this path, so a cloud read can happen on the thread that asked for it.
    /// </summary>
    internal sealed class GraphMailReader : IMailReader
    {
        /// <summary>The reason key for "there is no usable sign-in for this mailbox".</summary>
        internal const string NotSignedInReason = "error.graph.not_signed_in";

        /// <summary>The reason key for "the service could not be reached at all".</summary>
        internal const string OfflineReason = "error.graph.offline";

        /// <summary>
        /// The reason key for "the mailbox answered, but not with mail". The same key the desktop
        /// reader uses for a mid-scan failure, so the message a user sees does not depend on which
        /// source failed.
        /// </summary>
        internal const string ReadFailedReason = "error.scan.failed";

        /// <summary>
        /// The ceiling a full scan walks, the same two thousand items the desktop reader stops at: a
        /// mailbox with a hundred thousand mails in range must not turn into a hundred thousand rows.
        /// </summary>
        private const int MaxItems = 2000;

        /// <summary>
        /// How many pages one folder will be paged through before the reader stops. Two thousand items
        /// at the page size asked for is twenty pages; the cap exists so a service that kept handing
        /// back a next link could not loop forever.
        /// </summary>
        private const int MaxPages = 24;

        /// <summary>
        /// How many messages one page asks for: big enough that a few days of mail arrive in one round
        /// trip, small enough that the first page is not long in coming.
        /// </summary>
        private const int PageSize = 50;

        /// <summary>How many messages the cheap probe reads. Enough to find the newest one.</summary>
        private const int ProbePageSize = 5;

        private readonly string _clientId;
        private readonly RefreshTokenStore _tokens;
        private readonly IGraphTransport _http;

        private IList<GraphFolderInfo> _folders;
        private bool _tokenRead;
        private string _refreshToken;
        private string _accessToken;
        private DateTime _accessTokenExpiresOn = DateTime.MinValue;

        /// <summary>
        /// Set once the service has refused the stored sign-in, so one scan cannot retry a rejected
        /// token page after page. Cleared only by signing in again, which builds a new reader.
        /// </summary>
        private bool _signInRefused;

        /// <summary>
        /// Set once a call has failed to leave this machine: a scan that keeps retrying a dead network
        /// once per folder is slower than the user expects and tells them nothing new.
        /// </summary>
        private bool _unreachable;

        internal GraphMailReader(string clientId, RefreshTokenStore tokens, IGraphTransport http)
        {
            if (!GraphApp.HasClientId(clientId))
            {
                // A reader built without a usable registration could only fail later, at the first
                // request, as "not signed in" -- which is not what is wrong. Refusing to build one
                // keeps the two apart.
                throw new ArgumentException("a client id shaped like a GUID is required", "clientId");
            }

            if (tokens == null)
            {
                throw new ArgumentNullException("tokens");
            }

            if (http == null)
            {
                throw new ArgumentNullException("http");
            }

            _clientId = clientId.Trim();
            _tokens = tokens;
            _http = http;
        }

        /// <summary>The registration this reader identifies itself with. Public information.</summary>
        internal string ClientId
        {
            get { return _clientId; }
        }

        /// <summary>
        /// True when this mailbox could be read right now: a sign-in is stored for it and nothing has
        /// gone wrong yet. Deliberately cheap -- it reads no network and renews no token -- because the
        /// caller asks before every scan; a token that turns out to be refused during the scan stops
        /// that scan with the same reason this returns.
        /// </summary>
        public bool IsAvailable(out string reasonKey)
        {
            reasonKey = null;

            if (_signInRefused)
            {
                reasonKey = NotSignedInReason;
                return false;
            }

            if (_unreachable)
            {
                reasonKey = OfflineReason;
                return false;
            }

            if (RefreshToken() == null)
            {
                // Either nobody has signed in yet, or the stored token cannot be read on this machine
                // (it is encrypted for the user who signed in). To the user these are the same thing:
                // this mailbox needs a sign-in.
                reasonKey = NotSignedInReason;
                return false;
            }

            return true;
        }

        /// <summary>
        /// Forgets everything cached about this mailbox: the token, its expiry, the folder list and any
        /// failure so far. Called after signing in again (the stored token changed under us) and when
        /// the user asks for a fresh read. The next call re-reads all of it.
        /// </summary>
        internal void Refresh()
        {
            _tokenRead = false;
            _refreshToken = null;
            _accessToken = null;
            _accessTokenExpiresOn = DateTime.MinValue;
            _folders = null;
            _signInRefused = false;
            _unreachable = false;
        }

        /// <summary>
        /// The folder names this mailbox offers. A list that could not be read comes back empty: this
        /// member has no way to report why, and a scan -- which does -- reads the same list through
        /// Folders and reports the failure itself.
        /// </summary>
        public IList<string> ListFolders()
        {
            string failureKey;
            var folders = Folders(out failureKey);
            if (folders == null)
            {
                return new List<string>();
            }

            return GraphJson.FolderNames(folders) ?? new List<string>();
        }

        /// <summary>
        /// Reads the mail in range, up to the same two thousand item ceiling the desktop reader uses.
        ///
        /// A failure throws an InvalidOperationException carrying the reason key, which is the same
        /// channel the desktop reader uses, so the scan reports the reason instead of showing a short
        /// list the user would read as "no mail".
        /// </summary>
        public IEnumerable<MailSummary> Enumerate(MailScanScope scope, Action<int, int> progress, Func<bool> isCancelled)
        {
            var days = Math.Max(1, scope == null ? 30 : scope.Days);
            var cutoff = DateTime.Now.AddDays(-days);

            string failureKey;
            var folders = Folders(out failureKey);
            if (!string.IsNullOrEmpty(failureKey))
            {
                throw new InvalidOperationException(failureKey);
            }

            var targets = Targets(scope, folders);
            var rows = new List<MailSummary>();

            foreach (var target in targets)
            {
                if (isCancelled != null && isCancelled())
                {
                    break;
                }

                var url = MessagesUrl(target.Value, cutoff);
                var pages = 0;

                while (!string.IsNullOrEmpty(url) && pages < MaxPages && rows.Count < MaxItems)
                {
                    if (isCancelled != null && isCancelled())
                    {
                        break;
                    }

                    var token = AccessToken(out failureKey);
                    if (token == null)
                    {
                        throw new InvalidOperationException(failureKey ?? NotSignedInReason);
                    }

                    var response = _http.Get(url, token);
                    if (response == null || !response.IsSuccess)
                    {
                        throw new InvalidOperationException(RecordFailure(response, target.Key));
                    }

                    var page = GraphJson.ParseMessages(response.Body, target.Key);
                    if (page == null)
                    {
                        // A 2xx that is not a json page is a failed read, not an empty folder. This is
                        // the case the shared json guard exists for: an html page handed back by a
                        // proxy must never be reported as "no mail".
                        FileLogger.Info("Cloud mailbox answered with no readable page for " + target.Key);
                        throw new InvalidOperationException(ReadFailedReason);
                    }

                    foreach (var mail in page)
                    {
                        if (mail == null)
                        {
                            continue;
                        }

                        // The date filter is the service's; this is the reader's own check, so a page
                        // that ignores the filter cannot put an out-of-window mail in the list.
                        if (mail.ReceivedOn < cutoff)
                        {
                            continue;
                        }

                        rows.Add(mail);
                        if (rows.Count >= MaxItems)
                        {
                            break;
                        }
                    }

                    if (progress != null && page.Count > 0)
                    {
                        progress(rows.Count, MaxItems);
                    }

                    pages++;
                    url = GraphPage.NextLink(response.Body);
                }
            }

            return rows;
        }

        /// <summary>
        /// The cheap "did anything arrive?" answer: the head of the Inbox and its unread total, two
        /// small reads instead of a scan.
        ///
        /// Both halves are needed. The head catches new mail; the total catches the two cases the head
        /// cannot see -- mail a rule files straight out of the Inbox, and read/unread flips made in
        /// another client. A zero InboxCount means nothing was read, which the caller treats as "skip
        /// this tick" rather than "no change".
        /// </summary>
        public InboxProbe ProbeInbox()
        {
            var probe = new InboxProbe();

            string failureKey;
            var token = AccessToken(out failureKey);
            if (token == null)
            {
                return probe;
            }

            var head = _http.Get(HeadUrl(), token);
            if (head == null || !head.IsSuccess)
            {
                RecordFailure(head, "Inbox");
                return probe;
            }

            var newest = Newest(head.Body);
            var unread = InboxUnread(token);

            probe.Signature = (newest == null
                    ? "empty"
                    : InboxProbe.MakeFragment(newest.EntryId, newest.ReceivedOn))
                + "#" + (unread.HasValue
                    ? unread.Value.ToString(CultureInfo.InvariantCulture)
                    : "unknown");
            probe.UnreadCount = unread.HasValue ? unread.Value : 0;
            probe.InboxCount = 1;
            return probe;
        }

        /// <summary>
        /// False: a cloud mailbox has no Outlook window on this machine. The port says so on purpose,
        /// so the caller reports "cannot open here" rather than "mail not found".
        /// </summary>
        public bool TryOpenInOutlook(string entryId)
        {
            return false;
        }

        /// <summary>
        /// The refresh token, read from the encrypted store at most once per reader. A reader serves
        /// one signed-in mailbox and lives for one scan, so caching it here cannot outlive a sign-in.
        /// </summary>
        private string RefreshToken()
        {
            if (!_tokenRead)
            {
                _refreshToken = _tokens.Load();
                _tokenRead = true;
            }

            return _refreshToken;
        }

        /// <summary>
        /// An access token for the mailbox, renewed from the stored refresh token when the one in hand
        /// is close to expiring rather than after a call has already failed.
        ///
        /// Returns null and sets failureKey when there is no usable token; a caller that gets null has
        /// already had the reason recorded or read.
        /// </summary>
        private string AccessToken(out string failureKey)
        {
            failureKey = null;

            if (!string.IsNullOrEmpty(_accessToken) && !GraphAuth.NeedsRefresh(_accessTokenExpiresOn, DateTime.UtcNow))
            {
                return _accessToken;
            }

            var refresh = RefreshToken();
            if (refresh == null)
            {
                failureKey = NotSignedInReason;
                return null;
            }

            var response = _http.PostForm(GraphApp.TokenUrl, GraphAuth.RefreshForm(_clientId, refresh, GraphApp.ScopeText));
            if (response == null || !response.IsSuccess)
            {
                if (response == null || response.IsOffline)
                {
                    // The one failure the user can fix without touching the app.
                    failureKey = RecordFailure(response, null);
                    return null;
                }

                // Whatever else the service said, a sign-in cannot be renewed without a token, so the
                // answer to the user is the same: sign in again.
                _signInRefused = true;
                failureKey = NotSignedInReason;
                FileLogger.Info("Cloud sign-in could not be renewed ("
                    + response.StatusCode.ToString(CultureInfo.InvariantCulture) + "): " + ErrorCode(response));
                return null;
            }

            var token = GraphAuth.ParseToken(response.Body);
            if (token == null || !GraphAuth.HasToken(token))
            {
                _signInRefused = true;
                failureKey = NotSignedInReason;
                FileLogger.Info("Cloud sign-in answer carried no token: " + ErrorCode(response));
                return null;
            }

            _accessToken = token.AccessToken;
            _accessTokenExpiresOn = GraphAuth.ExpiresOn(DateTime.UtcNow, token.ExpiresInSeconds);
            StoreRotatedRefreshToken(token.RefreshToken);
            return _accessToken;
        }

        /// <summary>
        /// Saves a refresh token the service handed back with an access token. It is allowed to rotate
        /// the token, and keeping the new one is what lets a long-lived install stay signed in.
        ///
        /// A failure here is logged and swallowed: the read that just worked must not fail because a
        /// file could not be written.
        /// </summary>
        private void StoreRotatedRefreshToken(string refreshToken)
        {
            if (string.IsNullOrEmpty(refreshToken)
                || string.Equals(refreshToken, _refreshToken, StringComparison.Ordinal))
            {
                return;
            }

            try
            {
                _tokens.Save(refreshToken);
                _refreshToken = refreshToken;
                _tokenRead = true;
            }
            catch (Exception ex)
            {
                FileLogger.Error("Storing the renewed cloud sign-in failed", ex);
            }
        }

        /// <summary>
        /// The folder list, read once per reader: a scan asks for it before every folder it walks, and
        /// the list does not change under a read. Returns null and sets failureKey when the list could
        /// not be read -- a failed scan, not a mailbox with no folders in it.
        /// </summary>
        private IList<GraphFolderInfo> Folders(out string failureKey)
        {
            failureKey = null;

            if (_folders != null)
            {
                return _folders;
            }

            var token = AccessToken(out failureKey);
            if (token == null)
            {
                return null;
            }

            var response = _http.Get(GraphApp.FolderListUrl, token);
            if (response == null || !response.IsSuccess)
            {
                failureKey = RecordFailure(response, null);
                return null;
            }

            var folders = GraphJson.ParseFolders(response.Body);
            if (folders == null)
            {
                FileLogger.Info("Cloud mailbox answered with no readable folder list");
                failureKey = ReadFailedReason;
                return null;
            }

            _folders = folders;
            return _folders;
        }

        /// <summary>
        /// The folders this scan reads: the ones the scope names, or the Inbox when it names none this
        /// mailbox has.
        ///
        /// The desktop reader walks every folder of every store it can see. Over the network that would
        /// be hundreds of requests for a scan that looks at a few days, so a cloud scan reads the
        /// folders it was asked for, and falling back to the Inbox keeps a source that names a folder
        /// this mailbox calls something else from quietly reading nothing at all.
        /// </summary>
        private static IList<KeyValuePair<string, string>> Targets(MailScanScope scope, IList<GraphFolderInfo> folders)
        {
            var targets = new List<KeyValuePair<string, string>>();
            var wanted = scope == null ? null : scope.FolderPaths;

            if (wanted != null)
            {
                foreach (var name in wanted)
                {
                    if (string.IsNullOrEmpty(name))
                    {
                        continue;
                    }

                    var id = GraphJson.ResolveFolderId(folders, name);
                    if (!string.IsNullOrEmpty(id))
                    {
                        targets.Add(new KeyValuePair<string, string>(DisplayNameOf(folders, id, name), id));
                    }
                }
            }

            if (targets.Count == 0)
            {
                var id = GraphJson.ResolveFolderId(folders, "Inbox");
                if (string.IsNullOrEmpty(id))
                {
                    // The well-known name is itself a valid folder id, so this needs no folder list.
                    id = GraphJson.InboxWellKnownName;
                }

                targets.Add(new KeyValuePair<string, string>(DisplayNameOf(folders, id, "Inbox"), id));
            }

            return targets;
        }

        /// <summary>
        /// The mailbox's own name for a folder once its id is known: the list may call the Inbox
        /// something other than the English name, and the folder shown to the user should be the name
        /// the mailbox uses.
        /// </summary>
        private static string DisplayNameOf(IList<GraphFolderInfo> folders, string id, string fallback)
        {
            if (folders != null)
            {
                foreach (var folder in folders)
                {
                    if (folder != null && string.Equals(folder.Id, id, StringComparison.OrdinalIgnoreCase))
                    {
                        return string.IsNullOrEmpty(folder.DisplayName) ? fallback : folder.DisplayName;
                    }
                }
            }

            return fallback;
        }

        /// <summary>
        /// The url of the first page of one folder's mail in range, newest first.
        ///
        /// The range is the service's job: asking for every mail and dropping the old ones here would
        /// download a mailbox to look at a week of it. The reader checks the dates again on the way in,
        /// because a filter that is ignored must not become mail the app should not have shown.
        /// </summary>
        private static string MessagesUrl(string folderId, DateTime cutoff)
        {
            return GraphApp.MessagesUrl(folderId)
                + "?$top=" + PageSize.ToString(CultureInfo.InvariantCulture)
                + "&$select=" + GraphApp.SelectText
                + "&$orderby=" + Uri.EscapeDataString("receivedDateTime desc")
                + "&$filter=" + Uri.EscapeDataString("receivedDateTime ge " + GraphTime(cutoff));
        }

        /// <summary>
        /// The head of the Inbox, newest first. The well-known name is used rather than a folder id so
        /// the probe needs no folder list: it is the cheapest question the service can be asked.
        /// </summary>
        private static string HeadUrl()
        {
            return GraphApp.MessagesUrl(GraphJson.InboxWellKnownName)
                + "?$top=" + ProbePageSize.ToString(CultureInfo.InvariantCulture)
                + "&$select=" + GraphApp.SelectText
                + "&$orderby=" + Uri.EscapeDataString("receivedDateTime desc");
        }

        /// <summary>
        /// The Inbox asked for one field: how much of it is unread. The same well-known name, and a
        /// single field, so the probe stays two small reads.
        /// </summary>
        private static string UnreadCountUrl()
        {
            return GraphApp.MailFoldersUrl + "/" + Uri.EscapeDataString(GraphJson.InboxWellKnownName)
                + "?$select=unreadItemCount";
        }

        /// <summary>
        /// A moment written the way a Graph filter takes it: ISO 8601 in UTC, to the second. Rounding
        /// the sub-second part away keeps the comparison a comparison of whole seconds, which is what
        /// the service's date literals are.
        /// </summary>
        private static string GraphTime(DateTime value)
        {
            return value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        }

        /// <summary>The newest mail of one page, or null when the page holds none.</summary>
        private static MailSummary Newest(string body)
        {
            var page = GraphJson.ParseMessages(body, "Inbox");
            if (page == null)
            {
                return null;
            }

            MailSummary newest = null;
            foreach (var mail in page)
            {
                if (mail == null)
                {
                    continue;
                }

                if (newest == null || mail.ReceivedOn > newest.ReceivedOn)
                {
                    newest = mail;
                }
            }

            return newest;
        }

        /// <summary>
        /// The Inbox unread total, or null when it could not be read. Null is kept apart from zero: the
        /// count is part of the probe's signature, and reporting a failed read as "nothing unread"
        /// would hide mail.
        /// </summary>
        private int? InboxUnread(string token)
        {
            var response = _http.Get(UnreadCountUrl(), token);
            if (response == null || !response.IsSuccess)
            {
                FileLogger.Info("Cloud mailbox unread total could not be read");
                return null;
            }

            return GraphPage.UnreadCount(response.Body);
        }

        /// <summary>
        /// Turns one failed answer into the state the next call reads, one log line, and the reason key
        /// to report.
        ///
        /// The status and the folder are logged; the body is not, because a body is where a token would
        /// be, and a log file is exactly where a token must never appear.
        /// </summary>
        private string RecordFailure(GraphResponse response, string folder)
        {
            var where = string.IsNullOrEmpty(folder) ? string.Empty : " while reading " + folder;

            if (response == null || response.IsOffline)
            {
                _unreachable = true;
                FileLogger.Info("Cloud mailbox could not be reached" + where);
                return OfflineReason;
            }

            if (response.IsSignedOut)
            {
                _signInRefused = true;
                FileLogger.Info("Cloud mailbox refused the sign-in ("
                    + response.StatusCode.ToString(CultureInfo.InvariantCulture) + ")" + where);
                return NotSignedInReason;
            }

            FileLogger.Info("Cloud mailbox read failed ("
                + response.StatusCode.ToString(CultureInfo.InvariantCulture) + ")" + where);
            return ReadFailedReason;
        }

        /// <summary>
        /// The service's own error code for a failed call, for the log. Never the body: a token
        /// answer's body contains tokens, and a log file is exactly where they must not appear.
        /// </summary>
        private static string ErrorCode(GraphResponse response)
        {
            if (response == null || string.IsNullOrEmpty(response.Body))
            {
                return "no answer";
            }

            // The mailbox endpoints answer with an error object, the sign-in endpoint with a bare error
            // string; both shapes are tried so the log says why either failed.
            var graph = GraphJson.ParseErrorCode(response.Body);
            if (!string.IsNullOrEmpty(graph))
            {
                return graph;
            }

            var token = GraphAuth.ParseToken(response.Body);
            return token == null || string.IsNullOrEmpty(token.Error) ? "unknown error" : token.Error;
        }
    }
}
