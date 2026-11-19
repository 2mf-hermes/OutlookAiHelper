using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// v1.7.0: the scan range, and the only place that decides what the toolbar's day
    /// picker shows.
    ///
    /// The picker used to hold the five preset numbers and nothing else, so a stored range
    /// that was not one of them fell back to "30" the moment the toolbar was rebuilt. The
    /// range is now a number plus a "custom" entry: pick the entry and a validated sheet
    /// asks for the count, and the count itself is written into the picker's own list so
    /// the control reads back what is actually in force.
    ///
    /// The settings page edits the same number through <see cref="AppSettings.ScanDays"/>,
    /// and both surfaces go through <see cref="ScanDays"/>, so the toolbar and the scan
    /// that runs next cannot disagree about how far back to look.
    /// </summary>
    public partial class MainWindow
    {
        /// <summary>The toolbar entry that opens the custom-count sheet.</summary>
        private const string CustomDaysChoice = "days.custom.label";

        /// <summary>
        /// Guards the pickers while this class points them at a stored value. Assigning a
        /// selected index raises <see cref="SelectionChanged"/>, and a control being shown
        /// the range that is already in force must not be read as a fresh choice.
        /// </summary>
        private bool _suppressDaysUi;

        /// <summary>
        /// Which scan range is in force. Read from the settings object rather than cached:
        /// settings are loaded before the toolbar exists, so a field would hold a value from
        /// a previous construction.
        /// </summary>
        private int CurrentDays
        {
            get
            {
                int days;
                return _settings != null && ScanDays.TryParse(
                    _settings.ScanDays.ToString(CultureInfo.InvariantCulture), out days)
                        ? ScanDays.Clamp(days)
                        : ScanDays.Default;
            }
        }

        /// <summary>
        /// The toolbar's day picker: the presets, the range currently in force when it is not
        /// a preset, and the entry that opens the custom-count sheet.
        /// </summary>
        private ComboBox BuildDaysPicker()
        {
            var box = UiKit.Select();
            WithName(box, Strings.T("settings.days"));
            box.Width = 112;
            box.Height = 32;
            box.ToolTip = Strings.T("days.custom.body");

            foreach (var preset in ScanDays.Presets)
            {
                box.Items.Add(preset.ToString(CultureInfo.InvariantCulture));
            }

            box.Items.Add(Strings.T(CustomDaysChoice));
            box.SelectionChanged += OnDaysPickerChanged;
            return box;
        }

        /// <summary>
        /// Points the toolbar at a range without treating that as a choice, and brings the
        /// settings page's field into line so the two surfaces show the same number.
        /// </summary>
        private void SetDaysSelection(int days)
        {
            SyncDaysPickers();
        }

        /// <summary>
        /// Called when the toolbar picker's selection changes and after a settings save.
        /// Mirrors the range in force into every control that shows it.
        /// </summary>
        private void SyncDaysPickers()
        {
            var suppress = _suppressDaysUi;
            _suppressDaysUi = true;
            try
            {
                var inForce = CurrentDays;

                if (_daysBox != null)
                {
                    _daysBox.SelectionChanged -= OnDaysPickerChanged;
                    try
                    {
                        var text = inForce.ToString(CultureInfo.InvariantCulture);
                        var index = _daysBox.Items.IndexOf(text);
                        if (index < 0)
                        {
                            // A count the user typed is not one of the presets, so it has to be
                            // in the list to be shown. It is inserted rather than appended so the
                            // presets keep their order and the custom entry stays last.
                            _daysBox.Items.Insert(0, text);
                            index = 0;
                        }

                        _daysBox.SelectedIndex = index;
                    }
                    finally
                    {
                        _daysBox.SelectionChanged += OnDaysPickerChanged;
                    }
                }

                if (_settingsDays != null && !string.Equals(
                    _settingsDays.Text,
                    inForce.ToString(CultureInfo.InvariantCulture),
                    StringComparison.Ordinal))
                {
                    _settingsDays.Text = inForce.ToString(CultureInfo.InvariantCulture);
                }
            }
            finally
            {
                _suppressDaysUi = suppress;
            }
        }

        /// <summary>Selection changed in the toolbar picker.</summary>
        private void OnDaysPickerChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressDaysUi || _daysBox == null)
            {
                return;
            }

            var choice = _daysBox.SelectedItem as string;
            if (choice == null)
            {
                return;
            }

            if (choice == Strings.T(CustomDaysChoice))
            {
                // Put the previous range back first: the entry is a request to type a number,
                // not a range in itself, so the picker must not be left reading "custom".
                SyncDaysPickers();
                ShowCustomDaysSheet();
                return;
            }

            int days;
            if (int.TryParse(choice, NumberStyles.Integer, CultureInfo.InvariantCulture, out days))
            {
                OnScanDaysPicked(days, true);
            }
        }

        /// <summary>
        /// The custom-count sheet, reached from the toolbar. It writes the same setting the
        /// settings page's field writes, so the range the toolbar shows and the range the
        /// scan reports are one value rather than two that happen to agree today.
        /// </summary>
        private void ShowCustomDaysSheet()
        {
            if (!IsLoaded)
            {
                return;
            }

            var picked = CustomScanDaysSheet.Show(this, CurrentDays);
            if (picked.HasValue)
            {
                OnScanDaysPicked(picked.Value, true);
            }
        }

        /// <summary>
        /// Applies a range: it is clamped, written to the settings object, persisted, and
        /// mirrored to both surfaces. Called by the toolbar picker, by the custom sheet, and
        /// by the settings save.
        /// </summary>
        private void OnScanDaysPicked(int days, bool persist)
        {
            if (_suppressDaysUi)
            {
                return;
            }

            var clamped = ScanDays.Clamp(days);

            if (_settings == null)
            {
                _settings = AppSettings.CreateDefault();
            }

            if (_settings.ScanDays == clamped && !persist)
            {
                SyncDaysPickers();
                return;
            }

            _settings.ScanDays = clamped;

            if (persist)
            {
                try
                {
                    _settingsStore.Save(_settings);
                }
                catch (Exception ex)
                {
                    // A range that will not persist is still the range in force for this
                    // session, so the failure is logged and the status bar says so rather
                    // than the toolbar silently reverting.
                    Adapters.Logging.FileLogger.Error("OnScanDaysPicked", ex);
                    _statusText.Text = Strings.T("settings.days.invalid");
                    return;
                }
            }

            SyncDaysPickers();

            if (_statusText != null)
            {
                _statusText.Text = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.T("days.custom.saved"),
                    clamped);
            }
        }
    }
}