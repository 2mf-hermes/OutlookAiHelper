using System;
using OutlookAiHelper.Adapters.Mail;
using OutlookAiHelper.Application;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// The words a sign-in can end on, and the sentences the settings page shows for them.
    ///
    /// What is pinned here is that no known word reaches the screen as a key. Strings.T falls back
    /// to the key it was given, so a mistyped key would show a user "source.reason.offline" where a
    /// sentence belongs -- and a test is the only thing that would notice, because it compiles and
    /// runs perfectly well. Whether each key is translated in all three languages is the string
    /// audit's job: the audit is what reads every table.
    /// </summary>
    public static class SignInReasonTests
    {
        public static int Run()
        {
            var failed = 0;

            failed += Check("every word the sign-in can fail with has a sentence, and no sentence is a key", () =>
            {
                var words = new[]
                {
                    GraphSignIn.NoClientIdReason,
                    GraphSignIn.UnreadableReason,
                    GraphSignIn.OfflineReason,
                    GraphSignIn.NoCodeReason,
                    GraphSignIn.NoRefreshTokenReason,
                    ExpiredWord,
                    DeniedWord
                };

                foreach (var word in words)
                {
                    var shown = SignInReason.TextFor(word);
                    if (!IsASentence(shown, word))
                    {
                        Console.WriteLine("   word shown as a key or as itself: " + word + " -> " + shown);
                        return false;
                    }
                }

                return true;
            });

            failed += Check("the words that mean keep waiting read as waiting rather than as a failure", () =>
            {
                var waiting = SignInReason.TextFor(GraphAuth.Pending);

                // The two spellings the service uses have to end up in the same place, and the word
                // this test names has to be the word the sign-in actually reports -- going through
                // the constant is what ties the two together.
                return IsASentence(waiting, GraphAuth.Pending)
                    && SignInReason.TextFor(GraphAuth.SlowDown) == waiting
                    && SignInReason.TextFor(WaitingWord) == waiting;
            });

            failed += Check("a word the app has no sentence for is shown as itself", () =>
            {
                return SignInReason.TextFor("unauthorized_client") == "unauthorized_client"
                    && SignInReason.TextFor("http-403") == "http-403"
                    && !SignInReason.IsKnown("unauthorized_client")
                    && SignInReason.TextFor(string.Empty) == string.Empty;
            });

            failed += Check("every word this app can produce is one it knows, and nothing else claims to be", () =>
            {
                return SignInReason.IsKnown(GraphSignIn.NoClientIdReason)
                    && SignInReason.IsKnown(GraphSignIn.UnreadableReason)
                    && SignInReason.IsKnown(GraphSignIn.OfflineReason)
                    && SignInReason.IsKnown(GraphSignIn.NoCodeReason)
                    && SignInReason.IsKnown(GraphSignIn.NoRefreshTokenReason)
                    && SignInReason.IsKnown(GraphAuth.Pending)
                    && SignInReason.IsKnown(GraphAuth.SlowDown)
                    && SignInReason.IsKnown(ExpiredWord)
                    && SignInReason.IsKnown(DeniedWord)
                    && !SignInReason.IsKnown(null);
            });

            return failed;
        }

        /// <summary>
        /// A sentence, as opposed to the two things a missing translation looks like: the key that
        /// was asked for, and the reason word itself coming back untouched.
        /// </summary>
        private static bool IsASentence(string shown, string word)
        {
            return !string.IsNullOrEmpty(shown)
                && !string.Equals(shown, word, StringComparison.Ordinal)
                && !string.Equals(shown, "source.reason." + word, StringComparison.Ordinal)
                && !string.Equals(shown, "source.signIn.waiting", StringComparison.Ordinal)
                && !string.Equals(shown, "source.signIn." + word, StringComparison.Ordinal);
        }

        private static int Check(string name, Func<bool> body)
        {
            bool ok;
            try
            {
                ok = body();
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL " + name);
                Console.WriteLine("   " + ex.GetType().Name + ": " + ex.Message);
                return 1;
            }

            Console.WriteLine((ok ? "PASS " : "FAIL ") + name);
            return ok ? 0 : 1;
        }

        /// <summary>
        /// The service's own word for a code that outlived its lifetime, and its word for a consent
        /// that was refused. Spelled out here as well as in SignInReason so that the two cannot
        /// drift apart unnoticed: if the mapping ever stops covering one of them, the check above
        /// fails on this test's word rather than passing on a word nobody uses.
        /// </summary>
        private const string ExpiredWord = "expired_token";

        private const string DeniedWord = "access_denied";

        /// <summary>The word reported while the browser step is still running.</summary>
        private const string WaitingWord = "authorization_pending";
    }
}
