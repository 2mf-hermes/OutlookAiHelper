using System;
using System.Runtime.Serialization;

namespace OutlookAiHelper.Adapters.Mail
{
    /// <summary>
    /// The device code sign-in, written as the parts it really is: the forms to post, the answers
    /// to read, and the rule for deciding whether to ask again.
    ///
    /// Nothing here opens a socket. The transport posts the forms this class builds and hands the
    /// body back, which is what lets the whole flow be checked offline -- including the parts that
    /// otherwise only show up in a real tenant, like being told to slow down or being refused.
    ///
    /// The flow, in the order it happens: post the device code form, show the user the code and
    /// the address, then post the token form every few seconds until the answer is a token, a
    /// refusal, or the code has expired. What comes back in the end is a refresh token, and that
    /// is the only value on this path worth protecting; it never appears in a message this class
    /// can produce.
    /// </summary>
    internal static class GraphAuth
    {
        /// <summary>Waiting for the user to finish signing in. Not an error.</summary>
        internal const string Pending = "authorization_pending";

        /// <summary>Polling too fast. Wait longer, then keep going; the sign-in is still alive.</summary>
        internal const string SlowDown = "slow_down";

        /// <summary>The user refused, or the tenant refused on their behalf.</summary>
        internal const string AccessDenied = "access_denied";

        /// <summary>The code timed out or was already used. Signing in again is the only way on.</summary>
        internal const string ExpiredToken = "expired_token";

        /// <summary>How long to wait between polls when the answer does not say.</summary>
        internal const int DefaultPollSeconds = 5;

        /// <summary>How much longer to wait when the service says the polling is too fast.</summary>
        internal const int SlowDownAddedSeconds = 5;

        /// <summary>The code lifetime to assume when the answer does not say (Microsoft's usual 15 minutes).</summary>
        internal const int DefaultCodeLifetimeSeconds = 900;

        /// <summary>
        /// How early to replace an access token. A scan makes several calls in a row, so a token
        /// that expires between two of them would turn into a failed read, and the failure would
        /// look like a mailbox problem rather than an expiring credential.
        /// </summary>
        internal const int RefreshMarginSeconds = 300;

        internal const string DeviceCodeGrant = "urn:ietf:params:oauth:grant-type:device_code";
        internal const string RefreshGrant = "refresh_token";

        /// <summary>
        /// Asks for a code to show the user. The scope list travels as one escaped string, which
        /// is why it is built by GraphApp rather than pasted together at the call site.
        /// </summary>
        internal static string DeviceCodeForm(string clientId, string scopeText)
        {
            Require(clientId, "clientId");
            Require(scopeText, "scopeText");

            return "client_id=" + Escape(clientId) + "&scope=" + Escape(scopeText);
        }

        /// <summary>Trades the code for tokens. The same body is posted once per poll until it stops pending.</summary>
        internal static string CodeForm(string clientId, string deviceCode)
        {
            Require(clientId, "clientId");
            Require(deviceCode, "deviceCode");

            return "grant_type=" + Escape(DeviceCodeGrant)
                 + "&client_id=" + Escape(clientId)
                 + "&device_code=" + Escape(deviceCode);
        }

        /// <summary>
        /// Gets a new access token from the stored refresh token, which is how a source stays
        /// signed in between runs without the user seeing anything.
        /// </summary>
        internal static string RefreshForm(string clientId, string refreshToken, string scopeText)
        {
            Require(clientId, "clientId");
            Require(refreshToken, "refreshToken");
            Require(scopeText, "scopeText");

            return "grant_type=" + Escape(RefreshGrant)
                 + "&client_id=" + Escape(clientId)
                 + "&refresh_token=" + Escape(refreshToken)
                 + "&scope=" + Escape(scopeText);
        }

        /// <summary>
        /// The code, the address to type it into, and how long it lasts. Null when the answer
        /// carries no code -- there would be nothing to show the user, so there is no sign-in.
        /// </summary>
        internal static DeviceCodeInfo ParseDeviceCode(string json)
        {
            var document = JsonBody.Read<DeviceCodeDocument>(json);
            if (document == null || string.IsNullOrEmpty(document.DeviceCode) || string.IsNullOrEmpty(document.UserCode))
            {
                return null;
            }

            return new DeviceCodeInfo
            {
                DeviceCode = document.DeviceCode,
                UserCode = document.UserCode,
                VerificationUrl = document.VerificationUrl == null ? string.Empty : document.VerificationUrl,
                Message = document.Message == null ? string.Empty : document.Message,
                PollSeconds = PollSeconds(document.Interval),
                LifetimeSeconds = document.ExpiresIn > 0 ? document.ExpiresIn : DefaultCodeLifetimeSeconds
            };
        }

        /// <summary>
        /// What the token endpoint answered. A refusal is an answer too, so this returns the error
        /// rather than null: the caller has the three states to tell apart -- token, still pending,
        /// refused -- and they are not all failures.
        /// </summary>
        internal static TokenInfo ParseToken(string json)
        {
            var document = JsonBody.Read<TokenDocument>(json);
            if (document == null)
            {
                return null;
            }

            return new TokenInfo
            {
                AccessToken = string.IsNullOrEmpty(document.AccessToken) ? null : document.AccessToken,

                // Null, not empty, when the answer carries no fresh refresh token. The caller has to
                // be able to tell "keep the one you have" from "you have none": overwriting a good
                // refresh token with an empty string would end a sign-in that was working.
                RefreshToken = string.IsNullOrEmpty(document.RefreshToken) ? null : document.RefreshToken,

                ExpiresInSeconds = document.ExpiresIn,
                Error = string.IsNullOrEmpty(document.Error) ? null : document.Error,
                ErrorDescription = document.ErrorDescription == null ? string.Empty : document.ErrorDescription
            };
        }

        /// <summary>True when the answer carries a token that can be used.</summary>
        internal static bool HasToken(TokenInfo token)
        {
            return token != null && !string.IsNullOrEmpty(token.AccessToken);
        }

        /// <summary>True when the answer means "keep polling": still waiting, or polling too fast.</summary>
        internal static bool IsPending(string error)
        {
            return string.Equals(error, Pending, StringComparison.OrdinalIgnoreCase)
                || string.Equals(error, SlowDown, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>True when the answer means "poll again, but more slowly".</summary>
        internal static bool IsSlowDown(string error)
        {
            return string.Equals(error, SlowDown, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// The wait before the next poll. A slow_down answer lengthens it by the amount the
        /// service asks for; anything else keeps the current wait, and a wait that is missing or
        /// nonsense falls back to the documented default rather than becoming a busy loop.
        /// </summary>
        internal static int NextPollSeconds(int currentSeconds, string error)
        {
            var seconds = PollSeconds(currentSeconds);
            return IsSlowDown(error) ? seconds + SlowDownAddedSeconds : seconds;
        }

        /// <summary>When the access token stops working, counted from when it was issued.</summary>
        internal static DateTime ExpiresOn(DateTime issuedAtUtc, int expiresInSeconds)
        {
            return issuedAtUtc.AddSeconds(expiresInSeconds > 0 ? expiresInSeconds : 0);
        }

        /// <summary>
        /// True when the access token is close enough to expiry that it should be replaced before
        /// the next call instead of after the call fails.
        /// </summary>
        internal static bool NeedsRefresh(DateTime expiresOnUtc, DateTime nowUtc)
        {
            return nowUtc >= expiresOnUtc.AddSeconds(-RefreshMarginSeconds);
        }

        private static int PollSeconds(int seconds)
        {
            return seconds > 0 ? seconds : DefaultPollSeconds;
        }

        private static string Escape(string value)
        {
            return Uri.EscapeDataString(value);
        }

        private static void Require(string value, string name)
        {
            if (string.IsNullOrEmpty(value))
            {
                // Posting a form that cannot be answered only wastes a round trip and hides which
                // value was missing. Nothing on this path ever legitimately lacks these values.
                throw new ArgumentException("a value is required to build this request: " + name, name);
            }
        }

        [DataContract]
        private sealed class DeviceCodeDocument
        {
            [DataMember(Name = "device_code")]
            public string DeviceCode { get; set; }

            [DataMember(Name = "user_code")]
            public string UserCode { get; set; }

            [DataMember(Name = "verification_uri")]
            public string VerificationUrl { get; set; }

            [DataMember(Name = "message")]
            public string Message { get; set; }

            [DataMember(Name = "expires_in")]
            public int ExpiresIn { get; set; }

            [DataMember(Name = "interval")]
            public int Interval { get; set; }
        }

        [DataContract]
        private sealed class TokenDocument
        {
            [DataMember(Name = "access_token")]
            public string AccessToken { get; set; }

            [DataMember(Name = "refresh_token")]
            public string RefreshToken { get; set; }

            [DataMember(Name = "expires_in")]
            public int ExpiresIn { get; set; }

            [DataMember(Name = "error")]
            public string Error { get; set; }

            [DataMember(Name = "error_description")]
            public string ErrorDescription { get; set; }
        }
    }

    /// <summary>What the user needs in order to sign in: the code, where to type it, and how long it lasts.</summary>
    internal sealed class DeviceCodeInfo
    {
        public string DeviceCode { get; set; }
        public string UserCode { get; set; }
        public string VerificationUrl { get; set; }
        public string Message { get; set; }
        public int PollSeconds { get; set; }
        public int LifetimeSeconds { get; set; }
    }

    /// <summary>
    /// One answer from the token endpoint, whether it carried a token or a reason it did not.
    /// RefreshToken is null when the answer did not include one, which means the stored one is
    /// still current.
    /// </summary>
    internal sealed class TokenInfo
    {
        public string AccessToken { get; set; }
        public string RefreshToken { get; set; }
        public int ExpiresInSeconds { get; set; }
        public string Error { get; set; }
        public string ErrorDescription { get; set; }
    }
}
