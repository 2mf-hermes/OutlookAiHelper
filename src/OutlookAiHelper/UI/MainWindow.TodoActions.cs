using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OutlookAiHelper.Adapters.Logging;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// The follow-up affordances that go beyond add / tick / delete: waiting and overdue
    /// state, an editable note, a due date, one-click open of the source mail, list
    /// sorting, bulk clearing of finished items, and the sidebar count badge.
    ///
    /// Kept beside MainWindow.TodoDetail.cs (the read-only preview) so MainWindow.cs stays
    /// readable. Everything here works on the stored item, so it keeps working when the
    /// source mail has already left the scan window.
    /// </summary>
    public sealed partial class MainWindow
    {
        /// <summary>Days after which a still-waiting follow-up turns amber.</summary>
        private const int LongWaitDays = 7;

        /// <summary>
        /// Day offsets behind the due picker, in the order the choices are listed. 1 means
        /// "by the end of today", 0 clears the date.
        /// </summary>
        private static readonly int[] DueOffsets = { 0, 1, 3, 7, 30 };

        private StackPanel _todoActions;
        private TextBlock _todoWaitText;
        private TextBlock _todoDueText;
        private TextBox _todoNoteInput;
        private Button _todoOpenMailButton;
        private ComboBox _todoDuePicker;
        private ComboBox _todoSortPicker;
        private Border _navTodoBadgeHost;
        private TextBlock _navTodoBadge;

        /// <summary>Follow-up the detail panel is currently driving (null when it shows a mail).</summary>
        private TodoItem _detailTodo;

        private TodoSortMode _todoSortMode = TodoSortMode.Manual;

        /// <summary>Guards the due picker while the panel resets it to its prompt.</summary>
        private bool _suppressTodoDueUi;

        // ------------------------------------------------------------------ panel

        /// <summary>
        /// Builds the follow-up block that sits in the right-hand detail panel. The block is
        /// filled in by <see cref="RenderTodoActions"/> and stays hidden while the panel is
        /// showing a plain mail.
        /// </summary>
        private StackPanel BuildTodoActionsPanel()
        {
            var panel = new StackPanel
            {
                Visibility = Visibility.Collapsed,
                Margin = new Thickness(0, 10, 0, 2)
            };

            _todoWaitText = new TextBlock
            {
                FontSize = UiKit.TypeSubhead,
                FontWeight = FontWeights.SemiBold,
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 20
            };
            panel.Children.Add(_todoWaitText);

            _todoDueText = new TextBlock
            {
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 8, 0, 6)
            };
            panel.Children.Add(_todoDueText);

            _todoDuePicker = UiKit.Select();
            _todoDuePicker.Width = 180;
            _todoDuePicker.HorizontalAlignment = HorizontalAlignment.Left;
            _todoDuePicker.ToolTip = Strings.T("todo.action.setDue");
            _todoDuePicker.Items.Add(Strings.T("todo.action.setDue"));
            foreach (var offset in DueOffsets)
            {
                _todoDuePicker.Items.Add(DueChoiceLabel(offset));
            }

            _todoDuePicker.SelectedIndex = 0;
            _todoDuePicker.SelectionChanged += (s, e) => OnTodoDueChanged();
            panel.Children.Add(_todoDuePicker);

            panel.Children.Add(UiKit.Caption(Strings.T("todo.detail.note")));
            _todoNoteInput = UiKit.Input(string.Empty);
            _todoNoteInput.AcceptsReturn = true;
            _todoNoteInput.TextWrapping = TextWrapping.Wrap;
            _todoNoteInput.Height = 64;
            _todoNoteInput.VerticalContentAlignment = VerticalAlignment.Top;
            _todoNoteInput.ToolTip = Strings.T("todo.note.placeholder");
            panel.Children.Add(_todoNoteInput);

            var save = UiKit.Secondary(Strings.T("todo.note.save"), (s, e) => SaveTodoNote());
            save.HorizontalAlignment = HorizontalAlignment.Left;
            save.Margin = new Thickness(0, 8, 0, 0);
            panel.Children.Add(save);

            _todoOpenMailButton = UiKit.Secondary(Strings.T("todo.action.openMail"), (s, e) => OpenTodoSourceMail(_detailTodo));
            _todoOpenMailButton.HorizontalAlignment = HorizontalAlignment.Stretch;
            _todoOpenMailButton.Margin = new Thickness(0, 8, 0, 0);
            panel.Children.Add(_todoOpenMailButton);

            return panel;
        }

        private static string DueChoiceLabel(int offset)
        {
            if (offset <= 0)
            {
                return Strings.T("todo.due.none");
            }

            return offset == 1
                ? Strings.T("todo.due.today")
                : string.Format(Strings.T("todo.due.in"), offset);
        }

        /// <summary>Fills the follow-up block for <paramref name="item"/> and shows it.</summary>
        private void RenderTodoActions(TodoItem item)
        {
            try
            {
                if (_todoActions == null || item == null)
                {
                    SetTodoActionsVisible(false);
                    return;
                }

                _detailTodo = item;
                var now = DateTime.UtcNow;

                if (_todoWaitText != null)
                {
                    var text = WaitingStatusText(item, now);
                    _todoWaitText.Text = text ?? string.Empty;
                    _todoWaitText.Foreground = WaitingStatusBrush(item, now);
                    _todoWaitText.Visibility = string.IsNullOrEmpty(text) ? Visibility.Collapsed : Visibility.Visible;
                }

                if (_todoDueText != null)
                {
                    _todoDueText.Text = Strings.T("todo.detail.due") + ": " + DescribeDue(item);
                }

                if (_todoNoteInput != null)
                {
                    _todoNoteInput.Text = item.Note ?? string.Empty;
                }

                if (_todoOpenMailButton != null)
                {
                    var hasMail = !string.IsNullOrEmpty(item.SourceEntryId);
                    _todoOpenMailButton.IsEnabled = hasMail;
                    _todoOpenMailButton.Visibility = hasMail ? Visibility.Visible : Visibility.Collapsed;
                }

                ResetDuePicker();
                SetTodoActionsVisible(true);
            }
            catch (Exception ex)
            {
                FileLogger.Error("RenderTodoActions", ex);
            }
        }

        /// <summary>
        /// Turns the follow-up block on or off. The block carries its own "open the
        /// original" action, so the mail-level button would be a duplicate while it is up.
        /// </summary>
        private void SetTodoActionsVisible(bool visible)
        {
            if (_todoActions != null)
            {
                _todoActions.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
            }

            if (!visible)
            {
                _detailTodo = null;
            }

            if (_openOutlookButton != null)
            {
                _openOutlookButton.Visibility = visible ? Visibility.Collapsed : Visibility.Visible;
            }
        }

        private void ResetDuePicker()
        {
            if (_todoDuePicker == null)
            {
                return;
            }

            _suppressTodoDueUi = true;
            try
            {
                // Back to the prompt: the picker is an action, the line above it is the value.
                _todoDuePicker.SelectedIndex = 0;
            }
            finally
            {
                _suppressTodoDueUi = false;
            }
        }

        // ------------------------------------------------------------------ due date

        private void OnTodoDueChanged()
        {
            try
            {
                if (_suppressTodoDueUi || _detailTodo == null || _todoDuePicker == null)
                {
                    return;
                }

                var index = _todoDuePicker.SelectedIndex - 1;
                if (index < 0 || index >= DueOffsets.Length)
                {
                    return;
                }

                _todos.SetDue(_detailTodo.Id, DueFromOffset(DueOffsets[index]));
                _statusText.Text = Strings.T("todo.detail.due") + ": " + DescribeDue(_detailTodo);
                RenderTodoActions(_detailTodo);
                BindTodos();
            }
            catch (Exception ex)
            {
                FileLogger.Error("OnTodoDueChanged", ex);
            }
        }

        /// <summary>
        /// The UTC instant a "N 天後到期" choice means: the end of that local day, so a
        /// due date is never "overdue" the moment it is set.
        /// </summary>
        private static DateTime? DueFromOffset(int offset)
        {
            if (offset <= 0)
            {
                return null;
            }

            var day = DateTime.Now.Date.AddDays(offset - 1);
            var end = day.AddDays(1).AddTicks(-1);
            return DateTime.SpecifyKind(end, DateTimeKind.Local).ToUniversalTime();
        }

        private static string DescribeDue(TodoItem item)
        {
            if (item == null || !item.DueOn.HasValue)
            {
                return Strings.T("todo.due.none");
            }

            return item.DueOn.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm");
        }

        // ------------------------------------------------------------------ actions

        private void SaveTodoNote()
        {
            try
            {
                if (_detailTodo == null || _todoNoteInput == null)
                {
                    return;
                }

                _todos.SetNote(_detailTodo.Id, _todoNoteInput.Text);
                _statusText.Text = Strings.T("todo.note.saved");

                // The note is part of the item, so the row's copy stays in sync.
                BindTodos();
            }
            catch (Exception ex)
            {
                FileLogger.Error("SaveTodoNote", ex);
            }
        }

        private void OpenTodoSourceMail(TodoItem item)
        {
            try
            {
                if (item == null || string.IsNullOrEmpty(item.SourceEntryId))
                {
                    return;
                }

                if (!_reader.TryOpenInOutlook(item.SourceEntryId))
                {
                    _statusText.Text = Strings.T("error.outlook.not_detected");
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("OpenTodoSourceMail", ex);
                _statusText.Text = Strings.T("error.blank.prevented");
            }
        }

        private void OnClearDone()
        {
            try
            {
                var done = _todos.Items.Count(i => i != null && i.Status == TodoStatus.Done);
                if (done == 0)
                {
                    _statusText.Text = string.Format(Strings.T("todo.cleared"), 0);
                    return;
                }

                if (!ConfirmClearDone(done))
                {
                    return;
                }

                var cleared = _todos.ClearDone();
                if (_detailTodo != null && _detailTodo.Status == TodoStatus.Done)
                {
                    // The panel was showing one of the items we just deleted.
                    _detailTodoId = null;
                    SetTodoActionsVisible(false);
                }

                BindTodos();
                _statusText.Text = string.Format(Strings.T("todo.cleared"), cleared);
            }
            catch (Exception ex)
            {
                FileLogger.Error("OnClearDone", ex);
            }
        }

        /// <summary>Inline sheet, matching the update/restart dialogs - no stock MessageBox.</summary>
        private bool ConfirmClearDone(int count)
        {
            var dialog = new Window
            {
                Title = Strings.T("todo.action.clearDone"),
                Width = 420,
                Height = 230,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize,
                Background = Brushes.Transparent,
                FontFamily = Theme.UiFont
            };

            var stack = new StackPanel { Margin = new Thickness(24) };
            stack.Children.Add(new TextBlock
            {
                Text = string.Format(Strings.T("todo.clearDone.confirm"), count),
                FontSize = UiKit.TypeBody,
                Foreground = Theme.InkBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 20)
            });

            var row = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var cancel = UiKit.Secondary(Strings.T("action.cancel"), (s, e) => { dialog.DialogResult = false; });
            var clear = UiKit.Destructive(Strings.T("todo.action.clearDone"), (s, e) => { dialog.DialogResult = true; });
            cancel.Margin = new Thickness(0, 0, 8, 0);
            row.Children.Add(cancel);
            row.Children.Add(clear);
            stack.Children.Add(row);

            dialog.Content = UiKit.GlassPlate(stack, UiKit.RadiusSheet, new Thickness(28, 22, 28, 24));
            return dialog.ShowDialog() == true;
        }

        // ------------------------------------------------------------------ list tools

        /// <summary>Sort picker plus bulk clear, shown under the filter on the follow-up page.</summary>
        private UIElement BuildTodoToolsRow()
        {
            _todoSortPicker = UiKit.Select();
            _todoSortPicker.Width = 160;
            _todoSortPicker.ToolTip = Strings.T("todo.sort");
            _todoSortPicker.Items.Add(Strings.T("todo.sort.manual"));
            _todoSortPicker.Items.Add(Strings.T("todo.sort.waiting"));
            _todoSortPicker.Items.Add(Strings.T("todo.sort.due"));
            _todoSortPicker.Items.Add(Strings.T("todo.sort.quadrant"));
            _todoSortPicker.Items.Add(Strings.T("todo.sort.updated"));
            _todoSortPicker.SelectedIndex = SortModeToIndex(_todoSortMode);
            _todoSortPicker.SelectionChanged += (s, e) =>
            {
                _todoSortMode = IndexToSortMode(_todoSortPicker.SelectedIndex);
                BindTodos();
            };

            var clear = UiKit.Secondary(Strings.T("todo.action.clearDone"), (s, e) => OnClearDone());
            clear.Margin = new Thickness(8, 0, 0, 0);

            var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 8, 0, 0) };
            row.Children.Add(_todoSortPicker);
            row.Children.Add(clear);
            return row;
        }

        private static int SortModeToIndex(TodoSortMode mode)
        {
            switch (mode)
            {
                case TodoSortMode.LongestWaiting: return 1;
                case TodoSortMode.DueSoonest: return 2;
                case TodoSortMode.Quadrant: return 3;
                case TodoSortMode.RecentlyUpdated: return 4;
                default: return 0;
            }
        }

        private static TodoSortMode IndexToSortMode(int index)
        {
            switch (index)
            {
                case 1: return TodoSortMode.LongestWaiting;
                case 2: return TodoSortMode.DueSoonest;
                case 3: return TodoSortMode.Quadrant;
                case 4: return TodoSortMode.RecentlyUpdated;
                default: return TodoSortMode.Manual;
            }
        }

        // ------------------------------------------------------------------ shared labels

        /// <summary>
        /// "Waiting N days" / "Overdue by N days" / "Added today" for an open follow-up,
        /// or null for a finished one (the strikethrough already says it is done).
        /// </summary>
        private static string WaitingStatusText(TodoItem item, DateTime nowUtc)
        {
            if (item == null || item.Status == TodoStatus.Done)
            {
                return null;
            }

            if (item.IsOverdue(nowUtc))
            {
                return string.Format(Strings.T("todo.overdue"), OverdueDays(item, nowUtc));
            }

            var days = item.WaitingDays(nowUtc);
            return days <= 0
                ? Strings.T("todo.waiting.today")
                : string.Format(Strings.T("todo.waiting"), days);
        }

        private static Brush WaitingStatusBrush(TodoItem item, DateTime nowUtc)
        {
            if (item == null || item.Status == TodoStatus.Done)
            {
                return Theme.MutedBrush;
            }

            if (item.IsOverdue(nowUtc))
            {
                return Theme.DangerBrush;
            }

            return item.WaitingDays(nowUtc) >= LongWaitDays ? Theme.WarningBrush : Theme.MutedBrush;
        }

        /// <summary>Whole days past the due instant; a single hour late still reads as one.</summary>
        private static int OverdueDays(TodoItem item, DateTime nowUtc)
        {
            if (item == null || !item.DueOn.HasValue)
            {
                return 0;
            }

            var days = (int)Math.Floor((nowUtc - item.DueOn.Value).TotalDays) + 1;
            return days < 1 ? 1 : days;
        }

        // ------------------------------------------------------------------ sidebar badge

        /// <summary>Sidebar entry content: label plus the unfinished-count badge.</summary>
        private UIElement BuildNavTodoContent()
        {
            var stack = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };

            // No explicit Foreground: the pill's colour is inherited, so the label turns
            // white with the accent-filled (selected) state - same as the plain entries.
            stack.Children.Add(new TextBlock
            {
                Text = Strings.T("nav.todo"),
                VerticalAlignment = VerticalAlignment.Center
            });

            _navTodoBadge = new TextBlock
            {
                FontSize = UiKit.TypeCaption,
                FontWeight = FontWeights.SemiBold,
                Foreground = Theme.AccentBrush,
                VerticalAlignment = VerticalAlignment.Center
            };
            _navTodoBadgeHost = new Border
            {
                Child = _navTodoBadge,
                CornerRadius = UiKit.R(9),
                Padding = new Thickness(6, 1, 6, 1),
                Margin = new Thickness(8, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Visibility = Visibility.Collapsed
            };
            stack.Children.Add(_navTodoBadgeHost);

            UpdateTodoBadge();
            return stack;
        }

        /// <summary>Recounts open follow-ups. Safe to call before the nav is built.</summary>
        private void UpdateTodoBadge()
        {
            try
            {
                if (_navTodoBadge == null || _navTodoBadgeHost == null || _todos == null)
                {
                    return;
                }

                var open = _todos.Items.Count(i => i != null && i.Status == TodoStatus.Open);
                _navTodoBadge.Text = open > 99 ? "99+" : open.ToString();
                _navTodoBadgeHost.Visibility = open > 0 ? Visibility.Visible : Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                FileLogger.Error("UpdateTodoBadge", ex);
            }
        }

        /// <summary>
        /// Keeps the badge legible in both nav states: a white chip on the accent-filled
        /// pill, an accent wash on the transparent one.
        /// </summary>
        private void ApplyNavBadgeStyle(bool selected)
        {
            if (_navTodoBadgeHost == null)
            {
                return;
            }

            _navTodoBadgeHost.Background = selected
                ? new SolidColorBrush(Color.FromArgb(0xE6, 0xFF, 0xFF, 0xFF))
                : Theme.BadgeTintBrush;
        }
    }
}
