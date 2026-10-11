using System;
using OutlookAiHelper.Adapters.Mail;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// The device code sign-in, checked without a network and without an account.
    ///
    /// What can be pinned offline is everything the flow decides: which form carries which value,
    /// what an answer means, whether to ask again, and how long to wait before doing it. What
    /// cannot be pinned here is what a real tenant answers -- whether the mailbox allows the
    /// permission, whether the tenant permits device code sign-in at all -- and that stays on the
    /// list of things only a signed-in account can settle.
    /// </summary>
    public static class GraphAuthTests
    {
        private const string ClientId = "11111111-2222-3333-4444-555555555555";

        public static int Run()
        {
            var failed = 0;

            failed += Check("the device code request asks for exactly the permission the app uses", () =>
            {
                var form = GraphAuth.DeviceCodeForm(ClientId, GraphApp.ScopeText);

                return form.StartsWith("client_id=" + ClientId, StringComparison.Ordinal)
                    && form.IndexOf("&scope=", StringComparison.Ordinal) > 0
                    && form.IndexOf("Mail.ReadBasic", StringComparison.Ordinal) > 0
                    && form.IndexOf(' ') < 0;
            });

            failed += Check("the code request repeats the grant type the code was issued for", () =>
            {
                var form = GraphAuth.CodeForm("client-1", "DEVICE-CODE-1");

                return form.StartsWith("grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Adevice_code", StringComparison.Ordinal)
                    && form.IndexOf("&client_id=client-1", StringComparison.Ordinal) > 0
                    && form.IndexOf("&device_code=DEVICE-CODE-1", StringComparison.Ordinal) > 0;
            });

            failed += Check("a refresh token is escaped before it is posted, not pasted in", () =>
            {
                const string token = "M.C5+abc/def=ghi";
                var form = GraphAuth.RefreshForm("client-1", token, "offline_access Mail.ReadBasic");

                return form.StartsWith("grant_type=refresh_token", StringComparison.Ordinal)
                    && form.IndexOf("&refresh_token=M.C5%2Babc%2Fdef%3Dghi", StringComparison.Ordinal) > 0
                    && form.IndexOf(token, StringComparison.Ordinal) < 0;
            });

            failed += Check("a request that could never be answered is refused instead of posted", () =>
            {
                return RefusedValue(() => GraphAuth.RefreshForm("client-1", "", "offline_access"))
                    && RefusedValue(() => GraphAuth.CodeForm("", "DEVICE-CODE-1"))
                    && RefusedValue(() => GraphAuth.DeviceCodeForm(ClientId, ""));
            });

            failed += Check("a code answer gives the sign-in screen everything it shows", () =>
            {
                var info = GraphAuth.ParseDeviceCode(
                    @"{""user_code"":""ABCD-1234"",""device_code"":""DEVICE-CODE-1"","
                    + @"""verification_uri"":""https://microsoft.com/devicelogin"","
                    + @"""message"":""Use a browser to sign in"",""expires_in"":900,""interval"":5}");

                return info != null
                    && info.UserCode == "ABCD-1234"
                    && info.DeviceCode == "DEVICE-CODE-1"
                    && info.VerificationUrl == "https://microsoft.com/devicelogin"
                    && info.Message == "Use a browser to sign in"
                    && info.PollSeconds == 5
                    && info.LifetimeSeconds == 900;
            });

            failed += Check("a code answer with no usable interval waits the default instead of spinning", () =>
            {
                var missing = GraphAuth.ParseDeviceCode(@"{""user_code"":""A"",""device_code"":""D""}");
                var zero = GraphAuth.ParseDeviceCode(@"{""user_code"":""A"",""device_code"":""D"",""interval"":0}");

                return missing != null && missing.PollSeconds == GraphAuth.DefaultPollSeconds
                    && zero != null && zero.PollSeconds == GraphAuth.DefaultPollSeconds;
            });

            failed += Check("a code answer with no lifetime gets the usual fifteen minutes", () =>
            {
                var info = GraphAuth.ParseDeviceCode(@"{""user_code"":""A"",""device_code"":""D""}");
                return info != null && info.LifetimeSeconds == GraphAuth.DefaultCodeLifetimeSeconds;
            });

            failed += Check("an answer with no code cannot be a sign-in", () =>
            {
                return GraphAuth.ParseDeviceCode(@"{""interval"":5}") == null
                    && GraphAuth.ParseDeviceCode(@"{""device_code"":""DEVICE-CODE-1""}") == null;
            });

            failed += Check("an html page where a code should be is a failed sign-in, not an empty one", () =>
            {
                return GraphAuth.ParseDeviceCode("<html><body>sign in</body></html>") == null
                    && GraphAuth.ParseDeviceCode(string.Empty) == null;
            });

            failed += Check("a token answer carries the token, its lifetime and the next refresh token", () =>
            {
                var token = GraphAuth.ParseToken(
                    @"{""token_type"":""Bearer"",""scope"":""Mail.ReadBasic"",""expires_in"":3600,"
                    + @"""access_token"":""ACCESS-1"",""refresh_token"":""REFRESH-1""}");

                return GraphAuth.HasToken(token)
                    && token.AccessToken == "ACCESS-1"
                    && token.RefreshToken == "REFRESH-1"
                    && token.ExpiresInSeconds == 3600
                    && token.Error == null;
            });

            failed += Check("an answer with no new refresh token keeps the stored one", () =>
            {
                var token = GraphAuth.ParseToken(@"{""expires_in"":3600,""access_token"":""ACCESS-2""}");

                return GraphAuth.HasToken(token)
                    && token.RefreshToken == null
                    && token.ExpiresInSeconds == 3600;
            });

            failed += Check("waiting for the user is not a failure", () =>
            {
                var token = GraphAuth.ParseToken(
                    @"{""error"":""authorization_pending"",""error_description"":""Waiting for the user to sign in.""}");

                return token != null
                    && !GraphAuth.HasToken(token)
                    && token.Error == "authorization_pending"
                    && token.ErrorDescription == "Waiting for the user to sign in."
                    && GraphAuth.IsPending(token.Error)
                    && !GraphAuth.IsSlowDown(token.Error);
            });

            failed += Check("being told to slow down lengthens the wait instead of ending the sign-in", () =>
            {
                var token = GraphAuth.ParseToken(@"{""error"":""slow_down""}");

                return token != null
                    && GraphAuth.IsPending(token.Error)
                    && GraphAuth.IsSlowDown(token.Error)
                    && GraphAuth.NextPollSeconds(5, token.Error) == 5 + GraphAuth.SlowDownAddedSeconds
                    && GraphAuth.NextPollSeconds(0, token.Error) == GraphAuth.DefaultPollSeconds + GraphAuth.SlowDownAddedSeconds
                    && GraphAuth.NextPollSeconds(5, "authorization_pending") == 5
                    && GraphAuth.NextPollSeconds(0, null) == GraphAuth.DefaultPollSeconds;
            });

            failed += Check("a refusal and an expired code both stop the polling", () =>
            {
                var denied = GraphAuth.ParseToken(@"{""error"":""access_denied""}");
                var expired = GraphAuth.ParseToken(@"{""error"":""expired_token""}");

                return denied != null && !GraphAuth.IsPending(denied.Error) && denied.Error == "access_denied"
                    && expired != null && !GraphAuth.IsPending(expired.Error) && expired.Error == "expired_token"
                    && !GraphAuth.IsPending(null);
            });

            failed += Check("a token is replaced before it expires, not after a call fails", () =>
            {
                var issued = new DateTime(2024, 5, 6, 7, 0, 0, DateTimeKind.Utc);
                var expires = GraphAuth.ExpiresOn(issued, 3600);

                return expires == issued.AddSeconds(3600)
                    && !GraphAuth.NeedsRefresh(expires, expires.AddSeconds(-GraphAuth.RefreshMarginSeconds - 1))
                    && GraphAuth.NeedsRefresh(expires, expires.AddSeconds(-GraphAuth.RefreshMarginSeconds))
                    && GraphAuth.NeedsRefresh(expires, expires);
            });

            failed += Check("a truncated or non-object token answer is a failed answer, not a crash", () =>
            {
                return GraphAuth.ParseToken("{") == null
                    && GraphAuth.ParseToken("[]") == null
                    && GraphAuth.ParseToken("<html></html>") == null;
            });

            return failed;
        }

        private static bool RefusedValue(Func<string> body)
        {
            try
            {
                body();
                return false;
            }
            catch (ArgumentException)
            {
                return true;
            }
            catch (Exception)
            {
                return false;
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
