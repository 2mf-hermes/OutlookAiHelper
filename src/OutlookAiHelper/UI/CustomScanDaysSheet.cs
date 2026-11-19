using System;
using System.Globalization;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// The "custom" entry behind the toolbar's day picker: a single validated number that
    /// becomes the scan range. It shares <see cref="ScanDays"/> with the settings page, so
    /// no surface can store a count the reader would choke on.
    /// </summary>
    public sealed class CustomScanDaysSheet
    {
        private readonly Window _sheet;
        private readonly TextBox _input;
        private readonly TextBlock _error;

        public int? Value { get; private set; }

        private CustomScanDaysSheet(Window owner, int current)
        {
            _input = UiKit.Input(ScanDays.Clamp(current).ToString(CultureInfo.CurrentCulture));
            _input.Height = 38;
            _input.FontSize = UiKit.TypeHeadline;
            AutomationProperties.SetName(_input, Strings.T("days.custom.label"));
            _input.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.Key == Key.Enter)
                {
                    Accept();
                }
            };

            _error = new TextBlock
            {
                Text = Strings.T("days.custom.invalid"),
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.DangerBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(2, 6, 0, 0),
                Visibility = Visibility.Collapsed
            };

            var stack = new StackPanel();
            stack.Children.Add(new TextBlock
            {
                Text = Strings.T("days.custom.title"),
                FontSize = UiKit.TypeTitle2,
                FontWeight = FontWeights.SemiBold,
                Foreground = Theme.InkBrush,
                Margin = new Thickness(0, 0, 0, 4)
            });
            stack.Children.Add(new TextBlock
            {
                Text = Strings.T("days.custom.body"),
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 12)
            });
            stack.Children.Add(UiKit.Caption(Strings.T("days.custom.label")));
            stack.Children.Add(UiKit.FieldChrome(_input));
            stack.Children.Add(_error);

            var actions = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 18, 0, 0)
            };
            var cancel = UiKit.Secondary(Strings.T("todo.compose.cancel"), delegate { _sheet.Close(); });
            cancel.MinWidth = 88;
            cancel.Margin = new Thickness(0, 0, 10, 0);
            var ok = UiKit.Primary(Strings.T("days.custom.ok"), delegate { Accept(); });
            ok.MinWidth = 96;
            actions.Children.Add(cancel);
            actions.Children.Add(ok);

            var root = new DockPanel { LastChildFill = true };
            DockPanel.SetDock(actions, Dock.Bottom);
            root.Children.Add(actions);
            root.Children.Add(stack);

            _sheet = UiKit.GlassSheet(Strings.T("days.custom.title"), root, 440, 310);
            _sheet.Owner = owner;
            _sheet.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.Key == Key.Escape)
                {
                    _sheet.Close();
                }
            };
            _sheet.Loaded += delegate
            {
                _input.Focus();
                _input.SelectAll();
            };
        }

        /// <summary>Shows the sheet; returns the accepted count, or null when cancelled.</summary>
        public static int? Show(Window owner, int current)
        {
            var sheet = new CustomScanDaysSheet(owner, current);
            var accepted = sheet._sheet.ShowDialog();
            return accepted == true ? sheet.Value : (int?)null;
        }

        private void Accept()
        {
            int days;
            if (!ScanDays.TryParse(_input.Text, out days))
            {
                _error.Text = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.T("days.custom.invalid"),
                    ScanDays.Min,
                    ScanDays.Max);
                _error.Visibility = Visibility.Visible;
                _input.Focus();
                _input.SelectAll();
                return;
            }

            Value = days;
            _sheet.DialogResult = true;
        }
    }
}

