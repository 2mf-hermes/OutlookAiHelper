using System;
using System.Diagnostics;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OutlookAiHelper.Adapters.Update;
using OutlookAiHelper.Core.Application;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// The update flow, kept beside — not inside — the main window file:
    /// check → confirm → download → swap → restart. Steps that touch the network or the
    /// file system run off the UI thread, and none of them start without a click on a sheet.
    /// </summary>
    public sealed partial class MainWindow
    {
        private Button _updateCheckButton;
        private CheckBox _updateAutoCheck;
        private TextBlock _updateStatus;
        private bool _updateBusy;

        /// <summary>Settings section: the running build, a check button, and the opt-in toggle.</summary>
        private Panel BuildUpdateSection()
        {
            var panel = new StackPanel { Margin = new Thickness(0, 4, 0, 0) };
            panel.Children.Add(UiKit.Subtitle(Strings.T("update.section")));
            panel.Children.Add(new TextBlock
            {
                Text = Strings.T("update.current") + "   v" + AppInfo.Version,
                FontFamily = Theme.UiFont,
                FontSize = UiKit.TypeBody,
                Foreground = Theme.MutedBrush,
                Margin = new Thickness(0, 0, 0, 8)
            });

            var row = new StackPanel { Orientation = Orientation.Horizontal };
            _updateCheckButton = UiKit.Primary(Strings.T("update.check"), OnCheckUpdateClick);
            _updateCheckButton.MinWidth = 110;
            row.Children.Add(_updateCheckButton);

            _updateStatus = new TextBlock
            {
                FontFamily = Theme.UiFont,
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.MutedBrush,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(10, 0, 0, 0)
            };
            row.Children.Add(_updateStatus);
            panel.Children.Add(row);

            _updateAutoCheck = UiKit.AppleCheck(Strings.T("update.autostart"));
            _updateAutoCheck.IsChecked = _settings != null && _settings.CheckForUpdatesOnStartup;
            _updateAutoCheck.Margin = new Thickness(0, 12, 0, 4);
            panel.Children.Add(_updateAutoCheck);

            panel.Children.Add(new TextBlock
            {
                Text = Strings.T("update.security"),
                FontFamily = Theme.UiFont,
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 16)
            });

            return panel;
        }

        private async void OnCheckUpdateClick(object sender, RoutedEventArgs e)
        {
            if (_updateBusy)
            {
                return;
            }

            _updateBusy = true;
            SetUpdateStatus(Strings.T("update.checking"));
            if (_updateCheckButton != null)
            {
                _updateCheckButton.IsEnabled = false;
            }

            try
            {
                var offer = await Task.Run(() => CheckForUpdate());

                if (offer == null || offer.Status == UpdateStatus.Failed)
                {
                    SetUpdateStatus(string.Empty);
                    UpdateSheet.ShowFailure(this, offer == null ? "update.error.network" : offer.ErrorKey);
                    return;
                }

                if (offer.Status == UpdateStatus.NotPublished)
                {
                    SetUpdateStatus(string.Empty);
                    UpdateSheet.ShowNothingPublished(this);
                    return;
                }

                if (!offer.HasUpdate)
                {
                    SetUpdateStatus(string.Empty);
                    UpdateSheet.ShowUpToDate(this, offer.CurrentVersion);
                    return;
                }

                SetUpdateStatus(AvailableLabel(offer));
                var choice = UpdateSheet.ShowAvailable(this, offer);
                if (choice.Install)
                {
                    await InstallAsync(offer);
                }
                else if (choice.OpenPage)
                {
                    OpenReleasePage(offer);
                }
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("CheckUpdate", ex);
                SetUpdateStatus(string.Empty);
                UpdateSheet.ShowFailure(this, "update.error.network");
            }
            finally
            {
                _updateBusy = false;
                if (_updateCheckButton != null)
                {
                    _updateCheckButton.IsEnabled = true;
                }
            }
        }

        /// <summary>
        /// Silent, opt-in check at startup. It says nothing when there is nothing to say:
        /// a launch should never open with a network error the user did not ask for.
        /// </summary>
        private async Task StartupUpdateCheckAsync()
        {
            if (_settings == null || !_settings.CheckForUpdatesOnStartup)
            {
                return;
            }

            try
            {
                await Task.Delay(1500);
                var offer = await Task.Run(() => CheckForUpdate());
                if (offer == null || !offer.HasUpdate)
                {
                    return;
                }

                SetUpdateStatus(AvailableLabel(offer));
                var choice = UpdateSheet.ShowAvailable(this, offer);
                if (choice.Install)
                {
                    await InstallAsync(offer);
                }
                else if (choice.OpenPage)
                {
                    OpenReleasePage(offer);
                }
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("StartupUpdateCheck", ex);
            }
        }

        private static UpdateOffer CheckForUpdate()
        {
            return new CheckForUpdateUseCase(new GitHubReleaseFeed()).Execute(AppInfo.CurrentVersion());
        }

        private static string AvailableLabel(UpdateOffer offer)
        {
            return string.Format(CultureInfo.CurrentCulture, Strings.T("update.available.short"), offer.LatestVersion);
        }

        /// <summary>
        /// The only path that replaces the running program: download to %TEMP%, verify,
        /// then hand the swap to a detached script and exit. The write probe runs first so
        /// a read-only install folder is refused before anything is downloaded.
        /// </summary>
        private async Task InstallAsync(UpdateOffer offer)
        {
            string probeError;
            if (!UpdateInstaller.CanWriteToApplicationFolder(out probeError))
            {
                UpdateSheet.ShowFailure(this, probeError);
                return;
            }

            var progress = new UpdateProgressSheet(
                this,
                Strings.T("update.downloading.title"),
                Strings.T("update.downloading.body"));
            progress.Show();

            try
            {
                var installer = new UpdateInstaller();
                var errorKey = string.Empty;
                var downloaded = await Task.Run(() => installer.Download(
                    offer.AssetUrl,
                    offer.AssetName,
                    offer.AssetBytes,
                    percent => progress.Report(percent),
                    out errorKey));

                progress.Close();

                if (string.IsNullOrEmpty(downloaded))
                {
                    UpdateSheet.ShowFailure(this, errorKey);
                    return;
                }

                string swapError;
                if (!installer.TryStartSwap(downloaded, out swapError))
                {
                    UpdateSheet.ShowFailure(this, swapError);
                    return;
                }

                if (UpdateSheet.ConfirmRestart(this, offer.LatestVersion))
                {
                    System.Windows.Application.Current.Shutdown();
                }
                else
                {
                    SetUpdateStatus(string.Format(
                        CultureInfo.CurrentCulture, Strings.T("update.ready.body"), offer.LatestVersion));
                }
            }
            catch (Exception ex)
            {
                progress.Close();
                Adapters.Logging.FileLogger.Error("InstallUpdate", ex);
                UpdateSheet.ShowFailure(this, "update.error.download");
            }
        }

        private void OpenReleasePage(UpdateOffer offer)
        {
            var url = offer != null && !string.IsNullOrEmpty(offer.ReleasePageUrl)
                ? offer.ReleasePageUrl
                : AppInfo.ReleasePageUrl;

            try
            {
                Process.Start(url);
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("OpenReleasePage", ex);
            }
        }

        private void SetUpdateStatus(string text)
        {
            if (_updateStatus != null)
            {
                _updateStatus.Text = text ?? string.Empty;
            }
        }
    }
}
