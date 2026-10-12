using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using OutlookAiHelper.Adapters.Logging;
using OutlookAiHelper.Adapters.Mail;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// Signing a cloud source in: ask for a code, show it, then poll until the browser step is done.
    ///
    /// One attempt at a time -- a code on screen is a code somebody is typing -- so the attempt lives
    /// in fields rather than in the row. The rows are then drawn from those fields, which is what
    /// keeps a redraw from losing a code that is still valid. All of it lasts one run: a restart
    /// keeps the refresh token and nothing else.
    ///
    /// Every way this can fail comes back from the sign-in as a word, and SignInReason turns that
    /// word into a sentence. The word is the same one the reader would refuse a mailbox with, so the
    /// settings page and a scan never disagree about why a source is not being read.
    /// </summary>
    public sealed partial class MainWindow
    {
        private DispatcherTimer _signInTimer;
        private GraphSignIn _signIn;
        private DeviceCodeInfo _signInCode;
        private string _signInSourceId;
        private string _signInWord;
        private bool _signInBusy;

        /// <summary>
        /// The sign-in for one source, or null when this source has none to offer: a local source
        /// signs in through Outlook rather than through this app, and a build carrying no usable
        /// registration has nothing to sign in against.
        /// </summary>
        private static GraphSignIn SignInFor(MailSourceProfile source)
        {
            return MailReaderFactory.CreateSignIn(source);
        }

        /// <summary>
        /// Adds what signing in needs to one source's row: the code to type while an attempt is
        /// running, what became of the last one, and the button that starts or ends it.
        ///
        /// Left off entirely for a source with no sign-in to offer, rather than shown as a button
        /// whose only possible answer is a refusal.
        /// </summary>
        private void AddSignInControls(StackPanel text, StackPanel actions, MailSourceProfile source)
        {
            var signIn = SignInFor(source);
            if (signIn == null)
            {
                return;
            }

            var signedIn = signIn.IsSignedIn;
            var mine = _signInSourceId == source.Id;
            var waiting = mine && _signInCode != null;

            if (waiting)
            {
                // The address first and the code last and largest: the address is what the browser
                // needs, the code is what gets typed from this screen.
                text.Children.Add(SignInLine(Strings.T("source.signIn.hint"), Theme.MutedBrush, 12, false));
                text.Children.Add(SignInLine(_signInCode.VerificationUrl, Theme.MutedBrush, 12, false));
                text.Children.Add(SignInLine(_signInCode.UserCode, Theme.InkBrush, 16, true));
            }

            var status = StatusWord(signedIn, mine);
            if (status != null)
            {
                text.Children.Add(SignInLine(status, signedIn ? Theme.MutedBrush : Theme.WarningBrush, 12, false));
            }

            Button button;
            if (waiting)
            {
                button = UiKit.Secondary(Strings.T("source.signIn.cancel"), OnSignOutClick);
            }
            else if (signedIn)
            {
                button = UiKit.Secondary(Strings.T("source.signOut"), OnSignOutClick);
            }
            else
            {
                button = UiKit.Secondary(Strings.T("source.signIn"), OnSignInClick);
            }

            button.Tag = source.Id;
            button.Height = 32;
            button.MinWidth = 92;
            button.Margin = new Thickness(8, 0, 0, 0);
            actions.Children.Add(button);
        }

        /// <summary>
        /// What this row says about signing in, or null when it has nothing to say.
        ///
        /// A word from the attempt on this row is shown ahead of "signed in", because the attempt is
        /// why the row is being drawn again: a refusal that was immediately painted over with the
        /// older, calmer fact is a refusal nobody got to read.
        /// </summary>
        private string StatusWord(bool signedIn, bool mine)
        {
            if (mine && !string.IsNullOrEmpty(_signInWord))
            {
                return SignInReason.TextFor(_signInWord);
            }

            return signedIn ? Strings.T("source.signedIn") : null;
        }

        private static TextBlock SignInLine(string text, Brush brush, double size, bool strong)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = size,
                FontWeight = strong ? FontWeights.SemiBold : FontWeights.Normal,
                Foreground = brush,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
        }

        /// <summary>
        /// Starts a sign-in: ask for a code, show it on the row, then poll. Both network steps run
        /// off the UI thread, and a failed one arrives as a word rather than as an exception.
        /// </summary>
        private async void OnSignInClick(object sender, RoutedEventArgs e)
        {
            if (_signInBusy)
            {
                return;
            }

            var source = FindSourceById(TagOf(sender));
            var signIn = SignInFor(source);
            if (signIn == null)
            {
                return;
            }

            _signInBusy = true;
            _signIn = signIn;
            _signInSourceId = source.Id;
            _signInCode = null;
            _signInWord = null;
            StopSignInTimer();

            try
            {
                var code = await Task.Run(() => signIn.Begin());
                if (code == null)
                {
                    // Nothing to show: Reason says why, and StatusWord puts it on the row.
                    _signInWord = signIn.Reason;
                    FillSourceList();
                    return;
                }

                _signInCode = code;
                FillSourceList();
                StartSignInTimer(code.PollSeconds);
            }
            catch (Exception ex)
            {
                FileLogger.Error("SignIn", ex);
                _signInWord = GraphSignIn.UnreadableReason;
                _signInCode = null;
                FillSourceList();
            }
            finally
            {
                _signInBusy = false;
            }
        }

        /// <summary>
        /// Stops waiting and forgets the stored token, which is the whole of what "sign out" means
        /// here. The consent itself lives with Microsoft and is not touched: this only removes what
        /// this PC was keeping.
        /// </summary>
        private void OnSignOutClick(object sender, RoutedEventArgs e)
        {
            if (_signInBusy)
            {
                return;
            }

            var source = FindSourceById(TagOf(sender));
            var signIn = SignInFor(source);
            if (signIn == null)
            {
                return;
            }

            signIn.Forget();
            if (_signInSourceId == source.Id)
            {
                StopSignInTimer();
                _signIn = null;
                _signInCode = null;
                _signInWord = null;
                _signInSourceId = null;
            }

            // Signing out takes the stored token away, and the stored token is exactly what the
            // reader is refused by -- so a reader that is still alive for this source has to be let
            // go of. Without that, this run would keep reading on an access token that lives in
            // memory and nowhere else, and the app would be reading a mailbox it says it is signed
            // out of. Whether this is the mailbox being read decides how far the change reaches, the
            // same rule a finished sign-in follows.
            if (IsActiveSource(source.Id))
            {
                ApplySourceChoice();
            }
            else
            {
                FillSourceList();
            }
        }

        /// <summary>
        /// One poll. Still waiting leaves the code on screen and waits the interval the service asked
        /// for; anything else ends the attempt, which is what stops the timer.
        /// </summary>
        private async void OnSignInTimerTick(object sender, EventArgs e)
        {
            var signIn = _signIn;
            var code = _signInCode;
            if (_signInBusy || signIn == null || code == null)
            {
                return;
            }

            _signInBusy = true;
            try
            {
                var outcome = await Task.Run(() => signIn.Poll(code));
                if (outcome == SignInOutcome.Waiting)
                {
                    _signInWord = signIn.Reason;
                    _signInCode = code;
                    FillSourceList();
                    StartSignInTimer(signIn.PollSeconds);
                    return;
                }

                StopSignInTimer();
                _signInCode = null;
                _signInWord = outcome == SignInOutcome.SignedIn ? null : signIn.Reason;

                // Signing in hands the reader a token its next scan will use, so an attempt that
                // succeeded on the source being read rescans: the rows of mail may have been empty
                // or refused a minute ago, and a row saying "signed in" beside a list that still
                // belonged to nobody would be the wrong half of the story. A failure has nothing new
                // to read, so it only redraws.
                if (outcome == SignInOutcome.SignedIn && IsActiveSource(_signInSourceId))
                {
                    ApplySourceChoice();
                }
                else
                {
                    FillSourceList();
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("SignInPoll", ex);
                StopSignInTimer();
                _signInCode = null;
                _signInWord = GraphSignIn.UnreadableReason;
                FillSourceList();
            }
            finally
            {
                _signInBusy = false;
            }
        }

        private void StartSignInTimer(int seconds)
        {
            if (_signInTimer == null)
            {
                _signInTimer = new DispatcherTimer(DispatcherPriority.Background);
                _signInTimer.Tick += OnSignInTimerTick;
            }

            _signInTimer.Stop();
            _signInTimer.Interval = TimeSpan.FromSeconds(
                seconds > 0 ? seconds : GraphAuth.DefaultPollSeconds);
            _signInTimer.Start();
        }

        private void StopSignInTimer()
        {
            if (_signInTimer != null)
            {
                _signInTimer.Stop();
            }
        }

        /// <summary>The configured source with this id, or null when that row is already gone.</summary>
        private MailSourceProfile FindSourceById(string sourceId)
        {
            if (_settings == null || string.IsNullOrEmpty(sourceId))
            {
                return null;
            }

            return FindSource(_settings.MailSources, sourceId);
        }

        /// <summary>True when the source signed in is the one scans read.</summary>
        private bool IsActiveSource(string sourceId)
        {
            return _settings != null
                && !string.IsNullOrEmpty(sourceId)
                && string.Equals(_settings.ActiveMailSourceId, sourceId, StringComparison.Ordinal);
        }
    }
}
