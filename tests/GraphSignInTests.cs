using System;
using System.Collections.Generic;
using System.IO;
using OutlookAiHelper.Adapters.Mail;
using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Application;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// The device-code sign-in, checked without a network, an account, or a browser: which request
    /// each step sends, what each kind of answer does to the sign-in, and what ends up on disk.
    ///
    /// What cannot be pinned here -- whether a tenant allows the device-code flow at all, what its
    /// consent screen says, whether a Conditional Access policy refuses it -- needs a real account and
    /// stays on the open list.
    /// </summary>
    public static class GraphSignInTests
    {
        private const string ClientId = "11111111-2222-3333-4444-555555555555";
        private const string DeviceCodeText = "dc-1";
        private const string UserCodeText = "ABCD-1234";

        private static readonly List<string> TempFiles = new List<string>();

        public static int Run()
        {
            var failed = 0;

            try
            {
                failed += Check("a build with no client id cannot start a sign-in, and asks the service nothing", () =>
                {
                    var http = new FakeHttp();
                    var sign = new GraphSignIn(string.Empty, new RefreshTokenStore(TempPath()), http);

                    // Not one request: there is no registration to sign in against, so a round trip
                    // would only tell the app what it already knows.
                    return !sign.CanStart
                        && sign.Begin() == null
                        && sign.Reason == GraphSignIn.NoClientIdReason
                        && http.Urls.Count == 0;
                });

                failed += Check("a client id that is still a place-holder is not a client id", () =>
                {
                    var http = new FakeHttp();
                    var sign = new GraphSignIn("YOUR-CLIENT-ID-HERE", new RefreshTokenStore(TempPath()), http);

                    // Polling has to refuse for the same reason rather than let the missing client id
                    // turn into an exception on the UI thread.
                    var outcome = sign.Poll(new DeviceCodeInfo { DeviceCode = DeviceCodeText });

                    return !sign.CanStart
                        && sign.Begin() == null
                        && outcome == SignInOutcome.Failed
                        && sign.Reason == GraphSignIn.NoClientIdReason
                        && http.Urls.Count == 0;
                });

                failed += Check("asking for a code uses the device-code endpoint, for the read-only scopes only", () =>
                {
                    var http = Wired(url => Ok(DeviceCode()));
                    var sign = new GraphSignIn(ClientId, new RefreshTokenStore(TempPath()), http);

                    var code = sign.Begin();
                    if (code == null || http.Urls.Count != 1 || http.Forms.Count != 1)
                    {
                        return false;
                    }

                    var form = http.Forms[0];

                    // The scope text is the canary: were it missing from the form altogether, the two
                    // absences below would pass for the wrong reason.
                    return string.Equals(http.Urls[0], GraphApp.DeviceCodeUrl, StringComparison.Ordinal)
                        && form.IndexOf("client_id=" + ClientId, StringComparison.Ordinal) >= 0
                        && form.IndexOf(Uri.EscapeDataString(GraphApp.ScopeText), StringComparison.Ordinal) >= 0
                        && form.IndexOf("Mail.ReadBasic", StringComparison.Ordinal) >= 0
                        && form.IndexOf("Mail.ReadWrite", StringComparison.Ordinal) < 0
                        && form.IndexOf("Mail.Send", StringComparison.Ordinal) < 0
                        && code.DeviceCode == DeviceCodeText
                        && code.UserCode == UserCodeText
                        && code.PollSeconds == 5
                        && code.LifetimeSeconds == 900
                        && code.VerificationUrl.Length > 0;
                });

                failed += Check("a refused start keeps the service's own word for it", () =>
                {
                    var http = Wired(url => Status(400, "{\"error\":\"unauthorized_client\"}"));
                    var sign = new GraphSignIn(ClientId, new RefreshTokenStore(TempPath()), http);

                    return sign.Begin() == null && sign.Reason == "unauthorized_client";
                });

                failed += Check("a start that gets no answer is reported as the one failure the user can fix", () =>
                {
                    var http = Wired(url => Lost());
                    var sign = new GraphSignIn(ClientId, new RefreshTokenStore(TempPath()), http);

                    var code = sign.Begin();

                    return code == null && sign.Reason == GraphSignIn.OfflineReason;
                });

                failed += Check("a poll that is still pending arrives as a 400, which must not end the sign-in", () =>
                {
                    var http = Wired(TokenAnswer(400, "{\"error\":\"authorization_pending\"}"));
                    var sign = new GraphSignIn(ClientId, new RefreshTokenStore(TempPath()), http);

                    var code = sign.Begin();
                    var outcome = sign.Poll(code);

                    return code != null
                        && outcome == SignInOutcome.Waiting
                        && sign.Reason == GraphAuth.Pending
                        && sign.PollSeconds == GraphAuth.DefaultPollSeconds;
                });

                failed += Check("a poll told to slow down waits longer, and is still waiting", () =>
                {
                    var http = Wired(TokenAnswer(400, "{\"error\":\"slow_down\"}"));
                    var sign = new GraphSignIn(ClientId, new RefreshTokenStore(TempPath()), http);

                    var code = sign.Begin();
                    var outcome = sign.Poll(code);

                    return outcome == SignInOutcome.Waiting
                        && sign.Reason == GraphAuth.SlowDown
                        && sign.PollSeconds == GraphAuth.DefaultPollSeconds + GraphAuth.SlowDownAddedSeconds;
                });

                failed += Check("a finished sign-in stores the refresh token, and that file is what signed-in means", () =>
                {
                    var tokens = new RefreshTokenStore(TempPath());
                    var http = Wired(TokenAnswer(200, "{\"access_token\":\"AT-1\",\"refresh_token\":\"RT-1\",\"expires_in\":3600}"));
                    var sign = new GraphSignIn(ClientId, tokens, http);

                    var code = sign.Begin();
                    var outcome = sign.Poll(code);

                    // The poll belongs to the token endpoint and carries the code from the start; the
                    // refresh token is on disk, because the file -- not this object -- is what a later
                    // run reads.
                    return outcome == SignInOutcome.SignedIn
                        && sign.Reason == null
                        && sign.IsSignedIn
                        && tokens.Exists()
                        && tokens.Load() == "RT-1"
                        && http.Urls.Count == 2
                        && string.Equals(http.Urls[1], GraphApp.TokenUrl, StringComparison.Ordinal)
                        && http.Forms[1].IndexOf("device_code=" + DeviceCodeText, StringComparison.Ordinal) >= 0;
                });

                failed += Check("a token without a refresh token is not a sign-in, because it would not last", () =>
                {
                    var tokens = new RefreshTokenStore(TempPath());
                    var http = Wired(TokenAnswer(200, "{\"access_token\":\"AT-1\",\"expires_in\":3600}"));
                    var sign = new GraphSignIn(ClientId, tokens, http);

                    var outcome = sign.Poll(sign.Begin());

                    return outcome == SignInOutcome.Failed
                        && sign.Reason == GraphSignIn.NoRefreshTokenReason
                        && !tokens.Exists();
                });

                failed += Check("a declined sign-in is a failure with the service's own word", () =>
                {
                    var http = Wired(TokenAnswer(400, "{\"error\":\"access_denied\"}"));
                    var sign = new GraphSignIn(ClientId, new RefreshTokenStore(TempPath()), http);

                    var outcome = sign.Poll(sign.Begin());

                    return outcome == SignInOutcome.Failed
                        && sign.Reason == GraphAuth.AccessDenied
                        && sign.Reason != GraphAuth.Pending;
                });

                failed += Check("a code that ran out of time is a failure, not something to keep polling", () =>
                {
                    var http = Wired(TokenAnswer(400, "{\"error\":\"expired_token\"}"));
                    var sign = new GraphSignIn(ClientId, new RefreshTokenStore(TempPath()), http);

                    var outcome = sign.Poll(sign.Begin());

                    return outcome == SignInOutcome.Failed && sign.Reason == GraphAuth.ExpiredToken;
                });

                failed += Check("an answer it cannot read fails either way: no json on a 2xx, the status otherwise", () =>
                {
                    var html = Wired(TokenAnswer(200, "<html>not the token endpoint</html>"));
                    var unreadable = new GraphSignIn(ClientId, new RefreshTokenStore(TempPath()), html);
                    var first = unreadable.Poll(unreadable.Begin());

                    var broken = Wired(TokenAnswer(503, "<html>service unavailable</html>"));
                    var down = new GraphSignIn(ClientId, new RefreshTokenStore(TempPath()), broken);
                    var second = down.Poll(down.Begin());

                    // A 2xx that is not json is a failed step, never an empty result; a status the app
                    // cannot use is reported as itself rather than guessed at.
                    return first == SignInOutcome.Failed
                        && unreadable.Reason == GraphSignIn.UnreadableReason
                        && second == SignInOutcome.Failed
                        && down.Reason == "http-503";
                });

                failed += Check("signing out forgets the stored token, and doing it twice is harmless", () =>
                {
                    var tokens = new RefreshTokenStore(TempPath());
                    var sign = new GraphSignIn(ClientId, tokens, new FakeHttp());

                    tokens.Save("RT-1");
                    sign.Forget();
                    var forgotten = !tokens.Exists() && !sign.IsSignedIn;

                    // Nothing left to forget: this must not throw.
                    sign.Forget();

                    return forgotten && !tokens.Exists();
                });

                failed += Check("a poll with no code from a started sign-in fails without a request", () =>
                {
                    var http = new FakeHttp();
                    var sign = new GraphSignIn(ClientId, new RefreshTokenStore(TempPath()), http);

                    return sign.Poll(null) == SignInOutcome.Failed
                        && sign.Reason == GraphSignIn.NoCodeReason
                        && http.Urls.Count == 0;
                });

                failed += Check("both sign-in endpoints are the sign-in host, on the tenant that admits personal accounts", () =>
                {
                    const string Host = "https://login.microsoftonline.com/";
                    const string TenantPath = "/" + GraphApp.Tenant + "/";

                    // The tenant is what lets a personal outlook.com address sign in at all; a company
                    // tenant id here would quietly shut those accounts out.
                    return GraphApp.DeviceCodeUrl.IndexOf(Host, StringComparison.Ordinal) == 0
                        && GraphApp.TokenUrl.IndexOf(Host, StringComparison.Ordinal) == 0
                        && GraphApp.DeviceCodeUrl.IndexOf(TenantPath, StringComparison.Ordinal) > 0
                        && GraphApp.TokenUrl.IndexOf(TenantPath, StringComparison.Ordinal) > 0
                        && GraphApp.Tenant == "common";
                });
            }
            finally
            {
                Cleanup();
            }

            return failed;
        }

        private static FakeHttp Wired(Func<string, GraphResponse> onPost)
        {
            var http = new FakeHttp();
            http.OnPost = onPost;
            return http;
        }

        /// <summary>A transport that hands out a code when asked for one, and answers the poll as told.</summary>
        private static Func<string, GraphResponse> TokenAnswer(int status, string body)
        {
            return url => string.Equals(url, GraphApp.DeviceCodeUrl, StringComparison.Ordinal)
                ? Ok(DeviceCode())
                : Status(status, body);
        }

        private static string DeviceCode()
        {
            return "{\"device_code\":\"" + DeviceCodeText + "\""
                + ",\"user_code\":\"" + UserCodeText + "\""
                + ",\"verification_uri\":\"https://microsoft.com/devicelogin\""
                + ",\"message\":\"Enter the code " + UserCodeText + ".\""
                + ",\"expires_in\":900,\"interval\":5}";
        }

        private static GraphResponse Ok(string body)
        {
            return new GraphResponse { StatusCode = 200, Body = body };
        }

        private static GraphResponse Status(int statusCode, string body)
        {
            return new GraphResponse { StatusCode = statusCode, Body = body };
        }

        private static GraphResponse Lost()
        {
            // What the transport reports when nothing came back at all.
            return new GraphResponse { IsOffline = true };
        }

        private static string TempPath()
        {
            var path = Path.Combine(Path.GetTempPath(), "oah-p2a-" + Guid.NewGuid().ToString("N") + ".dat");
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

        private sealed class FakeHttp : IGraphTransport
        {
            private readonly List<string> _urls = new List<string>();
            private readonly List<string> _forms = new List<string>();

            internal Func<string, GraphResponse> OnPost;

            internal IList<string> Urls { get { return _urls; } }
            internal IList<string> Forms { get { return _forms; } }

            /// <summary>A sign-in reads nothing from the mailbox, so a GET here is a bug worth failing on.</summary>
            public GraphResponse Get(string url, string accessToken)
            {
                throw new InvalidOperationException("a sign-in must not read the mailbox: " + url);
            }

            public GraphResponse PostForm(string url, string form)
            {
                _urls.Add(url);
                _forms.Add(form == null ? string.Empty : form);
                return OnPost == null ? null : OnPost(url);
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
