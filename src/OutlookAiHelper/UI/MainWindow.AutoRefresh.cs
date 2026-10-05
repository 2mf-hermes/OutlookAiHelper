using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using OutlookAiHelper.Adapters.Logging;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// Automatic refresh: poll the head of the Inbox cheaply, and only pay for a full
    /// rescan when that probe actually saw something change.
    ///
    /// Rescanning on a plain timer would re-read every folder of every store whether or
    /// not anything happened — a lot of COM traffic on a mailbox, and completely
    /// invisible to the user when nothing changed. The probe keeps the list just as
    /// fresh at a fraction of the cost.
    ///
    /// One case a head-of-Inbox probe cannot see: mail a rule files straight into
    /// another folder never reaches the Inbox head. The safety ceiling below covers it.
    /// Both paths are read-only, like every other read from Outlook in this app.
    /// </summary>
    public sealed partial class MainWindow
    {
        /// <summary>Longest the list may go unrefreshed even if no probe saw a change.</summary>
        private static readonly TimeSpan AutoRefreshCeiling = TimeSpan.FromMinutes(30);

        /// <summary>
        /// Intervals the settings page offers, in minutes; 0 is off. Deliberately a short
        /// list — the probe is cheap, so "how often" is not a tuning problem the user
        /// should have to solve with a text box.
        /// </summary>
        private static readonly int[] AutoRefreshChoices = { 0, 1, 2, 5, 10, 15, 30 };

        private ComboBox _autoRefreshBox;
        private DispatcherTimer _autoRefreshTimer;
        private string _inboxSignature;
        private bool _autoRefreshBusy;
        private bool _scanRunning;
        private DateTime? _lastScanCompletedUtc;

        private static int AutoRefreshChoiceIndex(int minutes)
        {
            var index = Array.IndexOf(AutoRefreshChoices, minutes);
            return index >= 0 ? index : Array.IndexOf(AutoRefreshChoices, AppSettings.DefaultAutoRefreshMinutes);
        }

        private int SelectedAutoRefreshMinutes()
        {
            if (_autoRefreshBox == null)
            {
                return _settings != null ? _settings.AutoRefreshMinutes : AppSettings.DefaultAutoRefreshMinutes;
            }

            var index = _autoRefreshBox.SelectedIndex;
            if (index < 0 || index >= AutoRefreshChoices.Length)
            {
                return AppSettings.DefaultAutoRefreshMinutes;
            }

            return AutoRefreshChoices[index];
        }

        /// <summary>
        /// (Re)starts the probe timer from the saved settings. Called at startup and after
        /// the settings page is saved, so switching it off takes effect immediately.
        /// </summary>
        private void ConfigureAutoRefresh()
        {
            var minutes = _settings != null ? _settings.AutoRefreshMinutes : AppSettings.DefaultAutoRefreshMinutes;
            if (_autoRefreshTimer == null)
            {
                _autoRefreshTimer = new DispatcherTimer(DispatcherPriority.Background);
                _autoRefreshTimer.Tick += OnAutoRefreshTick;
            }

            _autoRefreshTimer.Stop();
            if (minutes <= 0)
            {
                FileLogger.Info("Auto refresh off");
                return;
            }

            _autoRefreshTimer.Interval = TimeSpan.FromMinutes(minutes);
            _autoRefreshTimer.Start();
            FileLogger.Info("Auto refresh every " + minutes + " minute(s)");
        }

        private async void OnAutoRefreshTick(object sender, EventArgs e)
        {
            try
            {
                await AutoRefreshTickAsync();
            }
            catch (Exception ex)
            {
                FileLogger.Error("AutoRefreshTick", ex);
            }
        }

        private async Task AutoRefreshTickAsync()
        {
            if (_autoRefreshBusy || _scanRunning || _settings == null || _settings.AutoRefreshMinutes <= 0)
            {
                return;
            }

            // A hidden or minimised window means nobody is reading the list: skip the COM
            // round trip instead of refreshing a surface that is not on screen.
            if (!IsVisible || WindowState == WindowState.Minimized)
            {
                return;
            }

            _autoRefreshBusy = true;
            try
            {
                var probe = await ProbeInboxAsync();
                if (probe == null || !probe.IsUsable)
                {
                    FileLogger.Info("Auto refresh probe unavailable; skipping this tick");
                    return;
                }

                var changed = !string.Equals(probe.Signature, _inboxSignature, StringComparison.Ordinal);
                var ceilingReached = !_lastScanCompletedUtc.HasValue
                    || DateTime.UtcNow - _lastScanCompletedUtc.Value >= AutoRefreshCeiling;
                if (!changed && !ceilingReached)
                {
                    return;
                }

                FileLogger.Info("Auto refresh rescan (changed=" + changed
                    + " ceiling=" + ceilingReached + " unread=" + probe.UnreadCount + ")");
                await RunScanAsync();
                _statusText.Text = string.Format(Strings.T("auto.refresh.updated"), _items.Count);
            }
            finally
            {
                _autoRefreshBusy = false;
            }
        }

        /// <summary>Reads the Inbox head on an STA thread; null when that is not possible.</summary>
        private async Task<InboxProbe> ProbeInboxAsync()
        {
            InboxProbe probe = null;
            await Task.Run(() =>
            {
                try
                {
                    // COM must run on STA — never Task.Run/MTA (that hard-terminates the process).
                    probe = Adapters.Com.ComSta.Run(() => _reader.ProbeInbox());
                }
                catch (Exception ex)
                {
                    FileLogger.Error("Inbox probe", ex);
                }
            });
            return probe;
        }

        /// <summary>
        /// Records the state the list on screen corresponds to. Called at the end of every
        /// scan, so the next probe compares against what is actually displayed — and a
        /// rescan cannot loop on its own side effects (mail shown as read, say).
        /// </summary>
        private async Task CaptureInboxSignatureAsync()
        {
            var probe = await ProbeInboxAsync();
            if (probe != null && probe.IsUsable)
            {
                _inboxSignature = probe.Signature;
            }

            _lastScanCompletedUtc = DateTime.UtcNow;
        }
    }
}
