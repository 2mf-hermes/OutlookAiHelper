using System;
using System.Globalization;
using OutlookAiHelper.Adapters.Mail;
using OutlookAiHelper.Adapters.Storage;

namespace OutlookAiHelper.Application
{
    /// <summary>What one poll of a device-code sign-in ended in.</summary>
    internal enum SignInOutcome
    {
        /// <summary>The browser step is not finished yet. Poll again after the current wait.</summary>
        Waiting,

        /// <summary>Done: a refresh token is stored, and the cloud source can be read from now on.</summary>
        SignedIn,

        /// <summary>Nothing more will come of this code, and <see cref="GraphSignIn.Reason"/> says what happened.</summary>
        Failed
    }

    /// <summary>
    /// The device-code sign-in, one step at a time: ask for a code to show the user, then poll until
    /// the answer is a token, a refusal, or the code dies.
    ///
    /// A screen drives this in a loop, so nothing here owns a timer or a clock -- the caller decides
    /// when to poll again (this class only says how long to wait), and stops when the code's own
    /// lifetime runs out. Nothing here holds user-visible text either: why a step failed comes back as
    /// <see cref="Reason"/>, a short word such as "authorization_pending", "access_denied" or
    /// "offline", and the screen turns that into a sentence in the reader's own language.
    ///
    /// Only a refresh token is ever kept, and only through <see cref="RefreshTokenStore"/>. An access
    /// token lives in memory for one run and is never written anywhere.
    /// </summary>
    internal sealed class GraphSignIn
    {
        /// <summary>Nothing to sign in against: no client id is configured, or what is there is not one.</summary>
        internal const string NoClientIdReason = "no-client-id";

        /// <summary>The answer was not the json this step needs.</summary>
        internal const string UnreadableReason = "unreadable";

        /// <summary>Nothing came back at all -- no network, or the host could not be reached.</summary>
        internal const string OfflineReason = "offline";

        /// <summary>Polling without a code from a started sign-in.</summary>
        internal const string NoCodeReason = "no-code";

        /// <summary>
        /// The answer carried an access token but no refresh token, so the sign-in would not outlive
        /// this run. Reported as a failure rather than as an hour-long success.
        /// </summary>
        internal const string NoRefreshTokenReason = "no-refresh-token";

        private readonly string _clientId;
        private readonly RefreshTokenStore _tokens;
        private readonly IGraphTransport _http;
        private int _pollSeconds = GraphAuth.DefaultPollSeconds;
        private string _reason;

        internal GraphSignIn(string clientId, RefreshTokenStore tokens, IGraphTransport http)
        {
            if (tokens == null) { throw new ArgumentNullException("tokens"); }
            if (http == null) { throw new ArgumentNullException("http"); }

            _clientId = clientId == null ? string.Empty : clientId;
            _tokens = tokens;
            _http = http;
        }

        /// <summary>
        /// False when there is no point starting: a build whose client id is empty or still a
        /// place-holder has no registration to sign in against, and must not spend a round trip
        /// finding out what it already knows.
        /// </summary>
        internal bool CanStart
        {
            get { return GraphApp.HasClientId(_clientId); }
        }

        /// <summary>True when a refresh token is stored for this source, which is what "signed in" means.</summary>
        internal bool IsSignedIn
        {
            get { return _tokens.Exists(); }
        }

        /// <summary>
        /// Why the last step ended the way it did: null when it signed in, a short word otherwise --
        /// including while it is still waiting, where that word is "authorization_pending".
        /// </summary>
        internal string Reason
        {
            get { return _reason; }
        }

        /// <summary>How long to wait before the next poll, as the service last asked.</summary>
        internal int PollSeconds
        {
            get { return _pollSeconds; }
        }

        /// <summary>
        /// The code to show, or null when there is nothing to show -- in which case Reason says why
        /// and the caller shows that instead. The wait the service asked for replaces the default.
        /// </summary>
        internal DeviceCodeInfo Begin()
        {
            _reason = null;

            if (!CanStart)
            {
                _reason = NoClientIdReason;
                return null;
            }

            var answer = _http.PostForm(GraphApp.DeviceCodeUrl, GraphAuth.DeviceCodeForm(_clientId, GraphApp.ScopeText));
            if (answer == null)
            {
                _reason = OfflineReason;
                return null;
            }

            // A code on a non-2xx would be as incoherent as a token on one: the body is read first,
            // but only a successful answer is allowed to be a sign-in.
            var code = GraphAuth.ParseDeviceCode(answer.Body);
            if (code != null)
            {
                if (!answer.IsSuccess)
                {
                    _reason = Why(answer);
                    return null;
                }

                _pollSeconds = code.PollSeconds;
                return code;
            }

            _reason = ServiceWord(answer);
            return null;
        }

        /// <summary>
        /// One poll. Waiting until the browser step is done, then either a stored refresh token and
        /// SignedIn, or Failed with Reason set.
        ///
        /// The order here matters: while the user is still signing in, the service refuses the poll
        /// with an error body and a 400, which is why a failed status cannot end a sign-in on its own.
        /// What the body says decides; the status only speaks when the body says nothing.
        /// </summary>
        internal SignInOutcome Poll(DeviceCodeInfo code)
        {
            _reason = null;

            if (code == null || string.IsNullOrEmpty(code.DeviceCode))
            {
                _reason = NoCodeReason;
                return SignInOutcome.Failed;
            }

            // Posting this form would be refused by GraphAuth for the same reason, but as an exception
            // on the UI thread. Refusing it here keeps the refusal a word the screen can show.
            if (!CanStart)
            {
                _reason = NoClientIdReason;
                return SignInOutcome.Failed;
            }

            var answer = _http.PostForm(GraphApp.TokenUrl, GraphAuth.CodeForm(_clientId, code.DeviceCode));
            if (answer == null)
            {
                _reason = OfflineReason;
                return SignInOutcome.Failed;
            }

            var token = GraphAuth.ParseToken(answer.Body);

            if (token != null && GraphAuth.HasToken(token))
            {
                if (!answer.IsSuccess)
                {
                    _reason = Why(answer);
                    return SignInOutcome.Failed;
                }

                if (string.IsNullOrEmpty(token.RefreshToken))
                {
                    // This app keeps nothing but the refresh token, and never stores an access token.
                    // Signing in on an access token alone would look fine for an hour and then stop,
                    // so it is reported as the failure it is.
                    _reason = NoRefreshTokenReason;
                    return SignInOutcome.Failed;
                }

                _tokens.Save(token.RefreshToken);
                return SignInOutcome.SignedIn;
            }

            if (token != null && !string.IsNullOrEmpty(token.Error))
            {
                var error = token.Error;

                // Slow down first: it also means "keep polling", but it is the one that changes the wait.
                if (GraphAuth.IsSlowDown(error))
                {
                    _pollSeconds = GraphAuth.NextPollSeconds(_pollSeconds, error);
                    _reason = GraphAuth.SlowDown;
                    return SignInOutcome.Waiting;
                }

                if (GraphAuth.IsPending(error))
                {
                    _reason = GraphAuth.Pending;
                    return SignInOutcome.Waiting;
                }

                // The service's own word for it -- "access_denied", "expired_token", and whatever else
                // it decides to answer -- is kept as it was sent rather than flattened into one failure.
                _reason = error;
                return SignInOutcome.Failed;
            }

            if (answer.IsOffline)
            {
                _reason = OfflineReason;
                return SignInOutcome.Failed;
            }

            // A 2xx that is not json is a failed step, never an empty result; a status the app cannot
            // use is reported as itself rather than guessed at.
            _reason = answer.IsSuccess ? UnreadableReason : Why(answer);
            return SignInOutcome.Failed;
        }

        /// <summary>
        /// Signs out: forgets the stored refresh token, so the next read asks for a sign-in again.
        /// Harmless when there was nothing stored.
        /// </summary>
        internal void Forget()
        {
            _reason = null;

            if (_tokens.Exists())
            {
                _tokens.Delete();
            }
        }

        /// <summary>What the answer said when it said nothing this step could use.</summary>
        private static string Why(GraphResponse answer)
        {
            if (answer == null || answer.IsOffline)
            {
                return OfflineReason;
            }

            return "http-" + answer.StatusCode.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// The refusal's own word from the body ("unauthorized_client" when the client id is wrong,
        /// usually), falling back to the status code when the body has none. The device-code and token
        /// answers carry that word in the same place, so the token reader is the one that can pick it
        /// out of either.
        /// </summary>
        private static string ServiceWord(GraphResponse answer)
        {
            var refusal = GraphAuth.ParseToken(answer.Body);
            if (refusal != null && !string.IsNullOrEmpty(refusal.Error))
            {
                return refusal.Error;
            }

            return Why(answer);
        }
    }
}
