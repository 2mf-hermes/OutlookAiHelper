using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using OutlookAiHelper.Adapters.Mail;
using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// The cloud reader, checked without a network and without an account: what it asks the service
    /// for, what it does with each kind of answer, and what it refuses to do.
    ///
    /// Everything here runs against a stand-in transport, so the decisions the app owns are pinned:
    /// which folders a scan reads, how far back it asks, how it follows a next link, when it stops,
    /// and which failures are reported as a failure rather than as an empty mailbox. What cannot be
    /// pinned here -- what a real tenant answers, whether its policy allows the device code sign-in,
    /// which fields Mail.ReadBasic actually returns -- needs a signed-in account and stays on the
    /// open list.
    /// </summary>
    public static class GraphReaderTests
    {
        private const string ClientId = "11111111-2222-3333-4444-555555555555";

        private static readonly List<string> TempFiles = new List<string>();

        public static int Run()
        {
            var failed = 0;

            try
            {
                failed += Check("a source with no stored sign-in says so instead of reading an empty mailbox", () =>
                {
                    var http = new FakeHttp();
                    http.OnGet = url => Ok(Page());
                    http.OnPost = url => Ok(Token());

                    var reader = new GraphMailReader(ClientId, new RefreshTokenStore(TempPath()), http);

                    string reason;
                    var available = reader.IsAvailable(out reason);

                    // Not one request: a source nobody has signed in to must not spend a round trip
                    // finding out what it already knows.
                    return !available
                        && reason == GraphMailReader.NotSignedInReason
                        && http.Gets.Count == 0
                        && http.Posts.Count == 0;
                });

                failed += Check("a stored sign-in is what makes a cloud source available", () =>
                {
                    var tokens = new RefreshTokenStore(TempPath());
                    var http = new FakeHttp();

                    string reason;
                    var before = new GraphMailReader(ClientId, tokens, http).IsAvailable(out reason);

                    tokens.Save("refresh-1");
                    var after = new GraphMailReader(ClientId, tokens, http).IsAvailable(out reason);

                    // The file, not the object, is what a source is signed in by: a reader made after
                    // the token is gone must not find one.
                    tokens.Delete();
                    var gone = new GraphMailReader(ClientId, tokens, http).IsAvailable(out reason);

                    return !before && after && !gone && reason == GraphMailReader.NotSignedInReason;
                });

                failed += Check("a scan asks for the folder it was given, for the window, newest first", () =>
                {
                    var http = new FakeHttp();
                    http.OnGet = url => Ok(FolderListUrl(url) ? Folders() : Page(Message("M-1", HoursAgo(1))));
                    http.OnPost = url => Ok(Token());

                    var reader = SignedIn(http);
                    var rows = Rows(reader, Names("Projects"), 7);

                    if (rows.Count != 1 || http.Gets.Count != 2)
                    {
                        return false;
                    }

                    if (!string.Equals(http.Gets[0], GraphApp.FolderListUrl, StringComparison.Ordinal))
                    {
                        return false;
                    }

                    var asked = http.Gets[1];
                    if (asked.IndexOf(GraphApp.MessagesUrl("f-2"), StringComparison.Ordinal) != 0)
                    {
                        return false;
                    }

                    // The url carries the window and the page size, and the filter names the moment
                    // the window starts -- give or take the seconds this test itself takes.
                    return asked.IndexOf("?$top=50&$select=", StringComparison.Ordinal) >= 0
                        && asked.IndexOf("&$orderby=" + Uri.EscapeDataString("receivedDateTime desc"), StringComparison.Ordinal) > 0
                        && FilterStartsAt(asked, DateTime.UtcNow.AddDays(-7))
                        && rows[0].EntryId == "M-1"

                        // The row is filed under the mailbox's own name for the folder, not under
                        // whatever the request happened to call it.
                        && rows[0].FolderPath == "Projects";
                });

                failed += Check("a scan that names no folder reads the Inbox, under the mailbox's own name for it", () =>
                {
                    var http = new FakeHttp();
                    http.OnGet = url => Ok(FolderListUrl(url) ? Folders() : Page(Message("M-1", HoursAgo(1))));
                    http.OnPost = url => Ok(Token());

                    var reader = SignedIn(http);
                    var rows = Rows(reader, Names("Inbox"), 7);

                    // The folder list calls the Inbox "Eingang"; the request called it "Inbox", which
                    // is matched on the well-known name, so the right folder is read...
                    if (rows.Count != 1
                        || http.Gets[1].IndexOf(GraphApp.MessagesUrl("f-1"), StringComparison.Ordinal) != 0)
                    {
                        return false;
                    }

                    // ...and shown the way the mailbox names it.
                    return rows[0].FolderPath == "Eingang";
                });

                failed += Check("a folder this mailbox does not have still reads the Inbox, not nothing", () =>
                {
                    var http = new FakeHttp();
                    http.OnGet = url => Ok(FolderListUrl(url) ? Folders() : Page(Message("M-1", HoursAgo(1))));
                    http.OnPost = url => Ok(Token());

                    var rows = Rows(SignedIn(http), Names("Nonexistent folder"), 7);

                    return rows.Count == 1
                        && http.Gets[1].IndexOf(GraphApp.MessagesUrl("f-1"), StringComparison.Ordinal) == 0;
                });

                failed += Check("a next link is followed, and its mail joins the same list", () =>
                {
                    var next = "https://graph.microsoft.com/v1.0/me/mailFolders/f-2/messages?$skiptoken=page-2";
                    var http = new FakeHttp();
                    http.OnGet = url =>
                    {
                        if (FolderListUrl(url))
                        {
                            return Ok(Folders());
                        }

                        return url.IndexOf("skiptoken", StringComparison.Ordinal) < 0
                            ? Ok(PageWithNext(next, Message("M-1", HoursAgo(1))))
                            : Ok(Page(Message("M-2", HoursAgo(2))));
                    };
                    http.OnPost = url => Ok(Token());

                    var rows = Rows(SignedIn(http), Names("Projects"), 7);

                    return rows.Count == 2
                        && http.Gets.Count == 3
                        && string.Equals(http.Gets[2], next, StringComparison.Ordinal)
                        && rows[0].EntryId == "M-1"
                        && rows[1].EntryId == "M-2";
                });

                failed += Check("a service that keeps handing back a next link is stopped, not followed forever", () =>
                {
                    var http = new FakeHttp();
                    http.OnGet = url =>
                    {
                        if (FolderListUrl(url))
                        {
                            return Ok(Folders());
                        }

                        // Always one more page: without a cap this never ends.
                        return Ok(PageWithNext(url, Message("M-1", HoursAgo(1))));
                    };
                    http.OnPost = url => Ok(Token());

                    var rows = Rows(SignedIn(http), Names("Projects"), 7);

                    // Bounded, and the bound was reached by reading pages, not by giving up early.
                    return http.Gets.Count <= 30 && rows.Count == http.Gets.Count - 1;
                });

                failed += Check("mail older than the window is dropped even when the page carries it", () =>
                {
                    var http = new FakeHttp();
                    http.OnGet = url => Ok(FolderListUrl(url)
                        ? Folders()
                        : Page(Message("M-old", HoursAgo(24 * 30)), Message("M-new", HoursAgo(1))));
                    http.OnPost = url => Ok(Token());

                    var rows = Rows(SignedIn(http), Names("Projects"), 7);

                    // The service was asked for the window; the reader checks it again, because a
                    // filter that is ignored must not become mail the app should not have shown.
                    return rows.Count == 1 && rows[0].EntryId == "M-new";
                });

                failed += Check("a refused sign-in stops the scan as a sign-in problem, and stays one", () =>
                {
                    var http = new FakeHttp();
                    http.OnGet = url => FolderListUrl(url) ? Ok(Folders()) : Refused();
                    http.OnPost = url => Ok(Token());

                    var reader = SignedIn(http);
                    var reason = Thrown(reader, Names("Projects"), 7);

                    if (reason != GraphMailReader.NotSignedInReason)
                    {
                        return false;
                    }

                    // The refusal is remembered, and availability is answered from it without a
                    // request: what the caller asks before every scan must stay cheap. The scan
                    // itself still reports the same reason rather than turning into "no mail".
                    var gets = http.Gets.Count;
                    string why;
                    var available = reader.IsAvailable(out why);

                    return !available
                        && why == GraphMailReader.NotSignedInReason
                        && http.Gets.Count == gets
                        && Thrown(reader, Names("Projects"), 7) == GraphMailReader.NotSignedInReason;
                });

                failed += Check("a service that cannot be reached is reported as offline, not as no mail", () =>
                {
                    var http = new FakeHttp();
                    http.OnGet = url => FolderListUrl(url) ? Ok(Folders()) : Lost();
                    http.OnPost = url => Ok(Token());

                    var reader = SignedIn(http);
                    var reason = Thrown(reader, Names("Projects"), 7);

                    string why;
                    var available = reader.IsAvailable(out why);

                    return reason == GraphMailReader.OfflineReason
                        && !available
                        && why == GraphMailReader.OfflineReason;
                });

                failed += Check("a 2xx that is not a json page is a failed read, not an empty folder", () =>
                {
                    var http = new FakeHttp();
                    http.OnGet = url => FolderListUrl(url)
                        ? Ok(Folders())
                        : Ok("<html><body>proxy says hello</body></html>");
                    http.OnPost = url => Ok(Token());

                    return Thrown(SignedIn(http), Names("Projects"), 7) == GraphMailReader.ReadFailedReason;
                });

                failed += Check("the probe reads the head of the Inbox and its unread total", () =>
                {
                    var http = new FakeHttp();
                    http.OnGet = url =>
                    {
                        if (url.IndexOf("unreadItemCount", StringComparison.Ordinal) > 0)
                        {
                            return Ok("{\"unreadItemCount\":4}");
                        }

                        return Ok(Page(Message("M-2", "2026-10-02T05:00:00Z"), Message("M-1", "2026-10-02T04:00:00Z")));
                    };
                    http.OnPost = url => Ok(Token());

                    var reader = SignedIn(http);
                    var probe = reader.ProbeInbox();
                    var expected = InboxProbe.MakeFragment("M-2", new DateTime(2026, 10, 2, 5, 0, 0, DateTimeKind.Utc));

                    // Two small reads instead of a scan, and the signature the timer compares.
                    return http.Gets.Count == 2
                        && probe.InboxCount == 1
                        && probe.UnreadCount == 4
                        && probe.Signature == expected + "#4"
                        && probe.IsUsable;
                });

                failed += Check("a probe whose Inbox could not be read reports nothing usable", () =>
                {
                    var http = new FakeHttp();
                    http.OnGet = url => Lost();
                    http.OnPost = url => Ok(Token());

                    var probe = SignedIn(http).ProbeInbox();

                    // Zero: the caller skips this tick rather than concluding "no change".
                    return probe.InboxCount == 0 && !probe.IsUsable;
                });

                failed += Check("cloud mail cannot be opened in Outlook, and says so by returning false", () =>
                {
                    var http = new FakeHttp();
                    return !SignedIn(http).TryOpenInOutlook("AAMkAD-1");
                });

                failed += Check("nothing a scan sends can change a mailbox, or ask for a message body", () =>
                {
                    var http = new FakeHttp();
                    http.OnGet = url => Ok(FolderListUrl(url) ? Folders() : Page(Message("M-1", HoursAgo(1))));
                    http.OnPost = url => Ok(Token());

                    Rows(SignedIn(http), Names("Projects"), 7);

                    // One POST, and it is the token renewal -- never a request against the mailbox.
                    if (http.Posts.Count != 1
                        || !http.Posts[0].StartsWith("grant_type=refresh_token", StringComparison.Ordinal))
                    {
                        return false;
                    }

                    // Every mailbox request went out as a GET, and carried the bearer token.
                    foreach (var token in http.Tokens)
                    {
                        if (token != "ACCESS-1")
                        {
                            return false;
                        }
                    }

                    // The read-only permission is what makes the excluded fields matter: the app never
                    // asks for a body, so the permission's exclusion costs it nothing.
                    return GraphApp.SelectText.IndexOf("body", StringComparison.OrdinalIgnoreCase) < 0
                        && GraphApp.SelectText.IndexOf("attachment", StringComparison.OrdinalIgnoreCase) < 0;
                });
            }
            finally
            {
                Cleanup();
            }

            return failed;
        }

        private static GraphMailReader SignedIn(FakeHttp http)
        {
            var tokens = new RefreshTokenStore(TempPath());
            tokens.Save("refresh-1");
            return new GraphMailReader(ClientId, tokens, http);
        }

        private static IList<string> Names(string name)
        {
            return new List<string> { name };
        }

        private static IList<MailSummary> Rows(GraphMailReader reader, IList<string> folders, int days)
        {
            var scope = new MailScanScope { FolderPaths = folders, Days = days };
            var rows = new List<MailSummary>();
            foreach (var mail in reader.Enumerate(scope, null, null))
            {
                rows.Add(mail);
            }

            return rows;
        }

        /// <summary>The reason key a scan threw, or null when it did not throw.</summary>
        private static string Thrown(GraphMailReader reader, IList<string> folders, int days)
        {
            try
            {
                Rows(reader, folders, days);
                return null;
            }
            catch (InvalidOperationException ex)
            {
                return ex.Message;
            }
        }

        private static bool FolderListUrl(string url)
        {
            return string.Equals(url, GraphApp.FolderListUrl, StringComparison.Ordinal);
        }

        private static bool FilterStartsAt(string url, DateTime expected)
        {
            const string marker = "&$filter=";
            var at = url.IndexOf(marker, StringComparison.Ordinal);
            if (at < 0)
            {
                return false;
            }

            const string prefix = "receivedDateTime ge ";
            var filter = Uri.UnescapeDataString(url.Substring(at + marker.Length));
            if (!filter.StartsWith(prefix, StringComparison.Ordinal))
            {
                return false;
            }

            DateTime cutoff;
            if (!DateTime.TryParse(filter.Substring(prefix.Length), CultureInfo.InvariantCulture,
                DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out cutoff))
            {
                return false;
            }

            return Math.Abs((cutoff - expected.ToUniversalTime()).TotalMinutes) < 5;
        }

        private static string HoursAgo(int hours)
        {
            return DateTime.UtcNow.AddHours(-hours).ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        }

        private static string TempPath()
        {
            var path = Path.Combine(Path.GetTempPath(), "oah-p1d-" + Guid.NewGuid().ToString("N") + ".dat");
            TempFiles.Add(path);
            return path;
        }

        private static void Cleanup()
        {
            foreach (var path in TempFiles)
            {
                try
                {
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                }
                catch (Exception)
                {
                    // A leftover temp file is not worth failing a test run over.
                }
            }

            TempFiles.Clear();
        }

        private static GraphResponse Ok(string body)
        {
            return new GraphResponse { StatusCode = 200, Body = body };
        }

        private static GraphResponse Refused()
        {
            return new GraphResponse
            {
                StatusCode = 401,
                Body = "{\"error\":{\"code\":\"InvalidAuthenticationToken\"}}"
            };
        }

        private static GraphResponse Lost()
        {
            // What the transport reports when nothing came back at all.
            return new GraphResponse { IsOffline = true };
        }

        private static string Token()
        {
            return "{\"access_token\":\"ACCESS-1\",\"expires_in\":3600}";
        }

        private static string Folders()
        {
            return "{\"value\":["
                + "{\"id\":\"f-1\",\"displayName\":\"Eingang\",\"wellKnownName\":\"inbox\"},"
                + "{\"id\":\"f-2\",\"displayName\":\"Projects\",\"wellKnownName\":null}"
                + "]}";
        }

        private static string Message(string id, string receivedOn)
        {
            return "{\"id\":\"" + id + "\",\"subject\":\"Subject " + id + "\""
                + ",\"from\":{\"emailAddress\":{\"name\":\"Ada Lovelace\",\"address\":\"ada@example.com\"}}"
                + ",\"receivedDateTime\":\"" + receivedOn + "\",\"isRead\":false}";
        }

        private static string Page(params string[] messages)
        {
            return "{\"value\":[" + string.Join(",", messages) + "]}";
        }

        private static string PageWithNext(string next, params string[] messages)
        {
            return "{\"value\":[" + string.Join(",", messages) + "]"
                + ",\"@odata.nextLink\":\"" + next + "\"}";
        }

        private sealed class FakeHttp : IGraphTransport
        {
            private readonly List<string> _gets = new List<string>();
            private readonly List<string> _posts = new List<string>();
            private readonly List<string> _tokens = new List<string>();

            internal Func<string, GraphResponse> OnGet;
            internal Func<string, GraphResponse> OnPost;

            internal IList<string> Gets { get { return _gets; } }
            internal IList<string> Posts { get { return _posts; } }
            internal IList<string> Tokens { get { return _tokens; } }

            public GraphResponse Get(string url, string accessToken)
            {
                _gets.Add(url);
                _tokens.Add(accessToken);
                return OnGet(url);
            }

            public GraphResponse PostForm(string url, string form)
            {
                _posts.Add(form ?? string.Empty);
                return OnPost(url);
            }
        }

        private static int Check(string name, Func<bool> body)
        {
            try
            {
                if (body())
                {
                    Console.WriteLine("PASS " + name);
                    return 0;
                }

                Console.WriteLine("FAIL " + name);
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL " + name + " :: " + ex.Message);
                return 1;
            }
        }
    }
}
