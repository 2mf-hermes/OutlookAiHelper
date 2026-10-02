using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OutlookAiHelper.Core;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    /// <summary>What the user chose on the "update available" sheet.</summary>
    internal sealed class UpdateChoice
    {
        public bool Install { get; set; }

        public bool OpenPage { get; set; }
    }

    /// <summary>
    /// macOS-style sheets for the update flow: a sheet reports one decision, the primary
    /// button sits rightmost, and the secondary is a quiet glass push. Everything the app
    /// is about to do — version, file, size, source — is stated on the sheet before the
    /// user is asked to agree to it.
    /// </summary>
    internal static class UpdateSheet
    {
        private const double SheetWidth = 460;
        private const double NotesHeight = 150;
        private const double ProgressTrackWidth = 380;

        public static UpdateChoice ShowAvailable(Window owner, UpdateOffer offer)
        {
            var choice = new UpdateChoice();
            var canInstall = offer != null && offer.CanAutoInstall;

            var dialog = CreateSheet(
                owner,
                Strings.T("update.title.available"),
                canInstall ? 560 : 480);

            var stack = new StackPanel { Margin = new Thickness(24) };

            stack.Children.Add(VersionHeadline(
                offer == null ? string.Empty : offer.CurrentVersion,
                offer == null ? string.Empty : offer.LatestVersion,
                offer != null && offer.Prerelease));

            stack.Children.Add(new TextBlock
            {
                Text = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.T("update.body.available"),
                    offer == null ? string.Empty : offer.CurrentVersion,
                    offer == null ? string.Empty : offer.LatestVersion),
                FontSize = UiKit.TypeBody,
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 14)
            });

            var notes = offer == null ? null : offer.Notes;
            if (!string.IsNullOrEmpty(notes) && notes.Trim().Length > 0)
            {
                stack.Children.Add(UiKit.Caption(Strings.T("update.notes")));
                var notesBox = new TextBox
                {
                    Text = notes.Trim(),
                    IsReadOnly = true,
                    TextWrapping = TextWrapping.Wrap,
                    AcceptsReturn = true,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    Background = Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    FontFamily = Theme.UiFont,
                    FontSize = UiKit.TypeFootnote,
                    Foreground = Theme.InkBrush,
                    Height = NotesHeight,
                    Padding = new Thickness(2),
                    IsTabStop = false
                };
                stack.Children.Add(UiKit.FieldChrome(notesBox));
                stack.Children.Add(new Border { Height = 14 });
            }

            // Provenance — shown before consent, not buried in a log.
            stack.Children.Add(MetaLine(
                Strings.T("update.source"),
                AppInfo.Owner + "/" + AppInfo.Repo + "  ·  GitHub Releases"));

            if (canInstall)
            {
                stack.Children.Add(MetaLine(Strings.T("update.file"), offer.AssetName));
                if (offer.AssetBytes > 0)
                {
                    stack.Children.Add(MetaLine(Strings.T("update.size"), FormatSize(offer.AssetBytes)));
                }
            }
            else if (offer != null && offer.AssetBlocked)
            {
                stack.Children.Add(MetaLine(Strings.T("update.file"), Strings.T("update.error.untrusted")));
            }

            stack.Children.Add(new TextBlock
            {
                Text = Strings.T("update.security"),
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 10, 0, 18)
            });

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            var later = UiKit.Secondary(Strings.T("update.later"), (s, e) => { dialog.DialogResult = false; });
            later.MinWidth = 88;
            later.Margin = new Thickness(0, 0, 8, 0);
            buttons.Children.Add(later);

            var primary = canInstall
                ? UiKit.Primary(Strings.T("update.install"), (s, e) =>
                {
                    choice.Install = true;
                    dialog.DialogResult = true;
                })
                : UiKit.Primary(Strings.T("update.page"), (s, e) =>
                {
                    choice.OpenPage = true;
                    dialog.DialogResult = true;
                });
            primary.MinWidth = 130;
            buttons.Children.Add(primary);

            stack.Children.Add(buttons);
            dialog.Content = UiKit.GlassPlate(stack, UiKit.RadiusSheet, new Thickness(28, 22, 28, 22));
            dialog.ShowDialog();
            return choice;
        }

        public static void ShowUpToDate(Window owner, string version)
        {
            var dialog = CreateSheet(owner, Strings.T("update.upToDate"), 400);
            var stack = new StackPanel { Margin = new Thickness(24) };
            stack.Children.Add(MessageBody(string.Format(
                CultureInfo.CurrentCulture, Strings.T("update.upToDate.body"), version)));

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            var ok = UiKit.Primary(Strings.T("update.ok"), (s, e) => { dialog.DialogResult = true; });
            ok.MinWidth = 88;
            buttons.Children.Add(ok);
            stack.Children.Add(buttons);

            dialog.Content = UiKit.GlassPlate(stack, UiKit.RadiusSheet, new Thickness(28, 22, 28, 22));
            dialog.ShowDialog();
        }

        public static void ShowNothingPublished(Window owner)
        {
            ShowInfo(owner, Strings.T("update.none.title"), Strings.T("update.none"));
        }

        public static void ShowFailure(Window owner, string errorKey)
        {
            var key = string.IsNullOrEmpty(errorKey) ? "update.error.download" : errorKey;
            ShowInfo(owner, Strings.T("update.failed.title"), Strings.T(key));
        }

        public static void ShowInfo(Window owner, string title, string body)
        {
            var dialog = CreateSheet(owner, title, 400);
            var stack = new StackPanel { Margin = new Thickness(24) };
            stack.Children.Add(MessageBody(body));

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            var ok = UiKit.Primary(Strings.T("update.ok"), (s, e) => { dialog.DialogResult = true; });
            ok.MinWidth = 88;
            buttons.Children.Add(ok);
            stack.Children.Add(buttons);

            dialog.Content = UiKit.GlassPlate(stack, UiKit.RadiusSheet, new Thickness(28, 22, 28, 22));
            dialog.ShowDialog();
        }

        public static bool ConfirmRestart(Window owner, string newVersion)
        {
            var dialog = CreateSheet(owner, Strings.T("update.ready.title"), 440);
            var stack = new StackPanel { Margin = new Thickness(24) };
            stack.Children.Add(MessageBody(string.Format(
                CultureInfo.CurrentCulture, Strings.T("update.ready.body"), newVersion)));

            var approved = false;
            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            var later = UiKit.Secondary(Strings.T("update.restart.later"), (s, e) => { dialog.DialogResult = false; });
            later.MinWidth = 150;
            later.Margin = new Thickness(0, 0, 8, 0);
            buttons.Children.Add(later);

            var now = UiKit.Primary(Strings.T("update.restart.now"), (s, e) =>
            {
                approved = true;
                dialog.DialogResult = true;
            });
            now.MinWidth = 140;
            buttons.Children.Add(now);
            stack.Children.Add(buttons);

            dialog.Content = UiKit.GlassPlate(stack, UiKit.RadiusSheet, new Thickness(28, 22, 28, 22));
            dialog.ShowDialog();
            return approved;
        }

        private static Window CreateSheet(Window owner, string title, double height)
        {
            return new Window
            {
                Title = title,
                Width = SheetWidth,
                Height = height,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = owner,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                Background = Brushes.Transparent
            };
        }

        private static UIElement VersionHeadline(string current, string latest, bool prerelease)
        {
            var row = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 10)
            };

            row.Children.Add(new TextBlock
            {
                Text = current,
                FontSize = UiKit.TypeTitle2,
                Foreground = Theme.MutedBrush,
                VerticalAlignment = VerticalAlignment.Center
            });
            row.Children.Add(new TextBlock
            {
                Text = "  →  ",
                FontSize = UiKit.TypeTitle2,
                Foreground = new SolidColorBrush(Theme.TertiaryLabel),
                VerticalAlignment = VerticalAlignment.Center
            });
            row.Children.Add(new TextBlock
            {
                Text = latest,
                FontSize = UiKit.TypeTitle2,
                FontWeight = FontWeights.SemiBold,
                Foreground = Theme.AccentBrush,
                VerticalAlignment = VerticalAlignment.Center
            });

            if (prerelease)
            {
                row.Children.Add(new TextBlock
                {
                    Text = "  " + Strings.T("update.prerelease"),
                    FontSize = UiKit.TypeCaption,
                    Foreground = new SolidColorBrush(Theme.Gold),
                    VerticalAlignment = VerticalAlignment.Bottom,
                    Margin = new Thickness(0, 0, 0, 3)
                });
            }

            return row;
        }

        private static UIElement MetaLine(string label, string value)
        {
            var row = new Grid { Margin = new Thickness(0, 0, 0, 4) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(64) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var caption = new TextBlock
            {
                Text = label,
                FontSize = UiKit.TypeCaption,
                Foreground = new SolidColorBrush(Theme.TertiaryLabel),
                VerticalAlignment = VerticalAlignment.Top
            };
            Grid.SetColumn(caption, 0);
            row.Children.Add(caption);

            var text = new TextBlock
            {
                Text = value,
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.InkBrush,
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetColumn(text, 1);
            row.Children.Add(text);

            return row;
        }

        private static TextBlock MessageBody(string body)
        {
            return new TextBlock
            {
                Text = body,
                FontSize = UiKit.TypeBody,
                Foreground = Theme.InkBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 20)
            };
        }

        internal static string FormatSize(long bytes)
        {
            if (bytes <= 0)
            {
                return Strings.T("update.size.unknown");
            }

            if (bytes >= 1048576L)
            {
                return (bytes / 1048576.0).ToString("0.0", CultureInfo.CurrentCulture) + " MB";
            }

            return Math.Max(1L, bytes / 1024L).ToString(CultureInfo.CurrentCulture) + " KB";
        }
    }

    /// <summary>
    /// Non-modal progress sheet. The download runs on a worker thread, so Report()
    /// marshals back to the window's dispatcher; the sheet refuses to be closed while
    /// bytes are still moving.
    /// </summary>
    internal sealed class UpdateProgressSheet
    {
        private readonly Window _window;
        private readonly Border _fill;
        private readonly TextBlock _status;
        private bool _allowClose;

        public UpdateProgressSheet(Window owner, string title, string body)
        {
            _window = new Window
            {
                Title = title,
                Width = 460,
                Height = 250,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = owner,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                Background = Brushes.Transparent
            };

            _window.Closing += (s, e) =>
            {
                if (!_allowClose)
                {
                    e.Cancel = true;
                }
            };

            var stack = new StackPanel { Margin = new Thickness(24) };
            stack.Children.Add(new TextBlock
            {
                Text = body,
                FontSize = UiKit.TypeBody,
                Foreground = Theme.InkBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 16)
            });

            var track = new Border
            {
                Width = 380,
                Height = 6,
                CornerRadius = UiKit.R(3),
                Background = new SolidColorBrush(Theme.Fill),
                HorizontalAlignment = HorizontalAlignment.Left
            };
            _fill = new Border
            {
                Width = 0,
                Height = 6,
                CornerRadius = UiKit.R(3),
                Background = Theme.AccentBrush,
                HorizontalAlignment = HorizontalAlignment.Left
            };
            var trackHost = new Grid { Width = 380, HorizontalAlignment = HorizontalAlignment.Left };
            trackHost.Children.Add(track);
            trackHost.Children.Add(_fill);
            stack.Children.Add(trackHost);

            _status = new TextBlock
            {
                Text = Strings.T("update.downloading.status"),
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.MutedBrush,
                Margin = new Thickness(0, 10, 0, 0)
            };
            stack.Children.Add(_status);

            _window.Content = UiKit.GlassPlate(stack, UiKit.RadiusSheet, new Thickness(28, 22, 28, 22));
        }

        public void Show()
        {
            _window.Show();
        }

        public void Report(int percent)
        {
            var action = new Action(delegate
            {
                var clamped = percent < 0 ? 0 : (percent > 100 ? 100 : percent);
                _fill.Width = 380.0 * clamped / 100.0;
                _status.Text = clamped + "%";
            });

            if (_window.Dispatcher.CheckAccess())
            {
                action();
            }
            else
            {
                _window.Dispatcher.BeginInvoke(action);
            }
        }

        public void Close()
        {
            _allowClose = true;
            try
            {
                _window.Close();
            }
            catch (Exception)
            {
            }
        }
    }
}
