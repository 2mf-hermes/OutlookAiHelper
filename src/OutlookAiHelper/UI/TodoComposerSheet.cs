using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// The sheet that collects everything a follow-up can carry: title and body (both
    /// quick-fillable from the source mail), a due day and a reminder, a repeat rule, a
    /// free-text category, files, and the source mail itself.
    ///
    /// The sheet owns no storage. It hands back a <see cref="TodoDraft"/>; the main
    /// window copies the picked files and saves the item. That split is what lets the
    /// draft be tested without a window, and keeps a failed file copy from leaving a
    /// half-written item behind.
    /// </summary>
    public sealed class TodoComposerSheet
    {
        private sealed class PendingFile
        {
            public string Path;
            public TodoAttachment Record;
        }

        private readonly Window _sheet;
        private readonly MailSummary _mail;
        private readonly TodoItem _existing;
        private readonly Border _fileHost;
        private readonly TextBlock _error;

        private readonly TextBox _title;
        private readonly TextBox _note;
        private readonly ComboBox _category;
        private readonly ComboBox _repeat;
        private readonly DatePicker _due;
        private readonly CheckBox _remind;
        private readonly DatePicker _remindDate;
        private readonly CheckBox _attachMail;

        private readonly List<PendingFile> _pending = new List<PendingFile>();
        private readonly List<string> _removed = new List<string>();

        public TodoDraft Result { get; private set; }

        private TodoComposerSheet(Window owner, MailSummary mail, TodoItem existing, IList<TodoAttachment> attachments, IList<string> categories, string prefillTitle)
        {
            _mail = mail;
            _existing = existing;

            var body = new StackPanel();

            // --- what the follow-up is -------------------------------------------
            body.Children.Add(SectionLabel(Strings.T("todo.compose.subject")));
            // A typed manual title is offered as the starting point, so routing a quick add
            // through this sheet costs the user nothing they had already written.
            _title = UiKit.Input(existing != null
                ? existing.Title
                : (string.IsNullOrWhiteSpace(prefillTitle)
                    ? (mail == null ? string.Empty : mail.Subject)
                    : prefillTitle));
            _title.Height = 38;
            AutomationProperties.SetName(_title, Strings.T("todo.compose.subject"));
            body.Children.Add(UiKit.FieldChrome(_title));

            if (mail != null)
            {
                var fillRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 6, 0, 0) };
                fillRow.Children.Add(SmallButton(Strings.T("todo.compose.fillSubject"), delegate
                {
                    if (!string.IsNullOrEmpty(mail.Subject))
                    {
                        _title.Text = mail.Subject;
                    }
                }));
                body.Children.Add(fillRow);
            }

            body.Children.Add(SectionLabel(Strings.T("todo.compose.note")));
            _note = UiKit.Input(existing == null ? string.Empty : (existing.Note ?? string.Empty));
            _note.AcceptsReturn = true;
            _note.TextWrapping = TextWrapping.Wrap;
            _note.Height = 96;
            _note.VerticalContentAlignment = VerticalAlignment.Top;
            _note.Padding = new Thickness(8, 6, 8, 6);
            _note.ToolTip = Strings.T("todo.compose.note.placeholder");
            AutomationProperties.SetName(_note, Strings.T("todo.compose.note"));
            body.Children.Add(UiKit.FieldChrome(_note));

            if (mail != null)
            {
                var noteRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 6, 0, 0) };
                noteRow.Children.Add(SmallButton(Strings.T("todo.compose.fillNote"), delegate { FillNoteFromMail(false); }));
                noteRow.Children.Add(SmallButton(Strings.T("todo.compose.fillFrom"), delegate { FillNoteFromMail(true); }));
                body.Children.Add(noteRow);
            }

            // --- when ------------------------------------------------------------
            body.Children.Add(SectionLabel(Strings.T("todo.compose.due")));
            _due = NewDatePicker(existing == null ? null : existing.DueOn, Strings.T("todo.compose.due"));
            body.Children.Add(UiKit.FieldChrome(_due));

            var quick = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 6, 0, 0) };
            quick.Children.Add(SmallButton(Strings.T("todo.compose.today"), delegate { SetDue(DateTime.Now.Date); }));
            quick.Children.Add(SmallButton(Strings.T("todo.compose.tomorrow"), delegate { SetDue(DateTime.Now.Date.AddDays(1)); }));
            quick.Children.Add(SmallButton(Strings.T("todo.compose.nextMonday"), delegate { SetDue(NextMonday()); }));
            quick.Children.Add(SmallButton(Strings.T("todo.compose.clearDate"), delegate
            {
                _due.SelectedDate = null;
                _remindDate.SelectedDate = null;
            }));
            body.Children.Add(quick);

            _remind = UiKit.AppleCheck(Strings.T("todo.compose.remind"));
            _remind.IsChecked = existing != null && existing.ReminderOn.HasValue;
            body.Children.Add(_remind);

            _remindDate = NewDatePicker(existing == null ? null : existing.ReminderOn, Strings.T("todo.compose.remindDate"));
            _remindDate.IsEnabled = _remind.IsChecked == true;
            _remind.Checked += delegate { _remindDate.IsEnabled = true; };
            _remind.Unchecked += delegate { _remindDate.IsEnabled = false; };
            _remindDate.SelectedDateChanged += delegate
            {
                // Picking a reminder day is the most direct way of saying "remind me", so
                // it turns the switch on rather than storing a date nothing would read.
                if (_remindDate.SelectedDate.HasValue && _remind.IsChecked != true)
                {
                    _remind.IsChecked = true;
                }
            };
            body.Children.Add(UiKit.FieldChrome(_remindDate));

            // --- how it comes back ------------------------------------------------
            body.Children.Add(SectionLabel(Strings.T("todo.compose.repeat")));
            _repeat = UiKit.Select();
            _repeat.Width = 200;
            _repeat.HorizontalAlignment = HorizontalAlignment.Left;
            _repeat.Items.Add(Strings.T("todo.repeat.none"));
            _repeat.Items.Add(Strings.T("todo.repeat.daily"));
            _repeat.Items.Add(Strings.T("todo.repeat.weekly"));
            _repeat.Items.Add(Strings.T("todo.repeat.monthly"));
            _repeat.Items.Add(Strings.T("todo.repeat.yearly"));
            _repeat.SelectedIndex = (int)(existing == null ? TodoRepeat.None : existing.Repeat);
            AutomationProperties.SetName(_repeat, Strings.T("todo.compose.repeat"));
            body.Children.Add(UiKit.FieldChrome(_repeat));

            // --- which pile -------------------------------------------------------
            body.Children.Add(SectionLabel(Strings.T("todo.compose.category")));
            _category = UiKit.Select();
            _category.IsEditable = true;
            _category.Width = 240;
            _category.HorizontalAlignment = HorizontalAlignment.Left;
            _category.ToolTip = Strings.T("todo.compose.category.placeholder");
            AutomationProperties.SetName(_category, Strings.T("todo.compose.category"));
            if (categories != null)
            {
                foreach (var name in categories)
                {
                    if (!string.IsNullOrEmpty(name))
                    {
                        _category.Items.Add(name);
                    }
                }
            }

            _category.Text = existing == null ? string.Empty : (existing.Category ?? string.Empty);
            body.Children.Add(UiKit.FieldChrome(_category));

            // --- the mail ---------------------------------------------------------
            body.Children.Add(SectionLabel(Strings.T("todo.compose.mail")));
            if (mail != null)
            {
                _attachMail = UiKit.AppleCheck(string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.T("todo.compose.mail.attach"),
                    ShortSubject()));
                _attachMail.IsChecked = true;
                body.Children.Add(_attachMail);
                body.Children.Add(new TextBlock
                {
                    Text = MailLine(mail),
                    FontSize = UiKit.TypeCaption,
                    Foreground = Theme.MutedBrush,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    Margin = new Thickness(2, 0, 0, 4)
                });
            }
            else
            {
                _attachMail = null;
                body.Children.Add(new TextBlock
                {
                    Text = existing != null && !string.IsNullOrEmpty(existing.SourceEntryId)
                        ? Strings.T("todo.compose.mail.keep")
                        : Strings.T("todo.compose.mail.none"),
                    FontSize = UiKit.TypeCaption,
                    Foreground = Theme.MutedBrush,
                    Margin = new Thickness(2, 0, 0, 4)
                });
            }

            // --- files ------------------------------------------------------------
            body.Children.Add(SectionLabel(Strings.T("todo.compose.files")));
            _fileHost = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0xF2, 0xF0, 0xEB)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xD0, 0xCB, 0xC3)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(10, 8, 10, 8),
                MinHeight = 44
            };
            body.Children.Add(_fileHost);

            var fileRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 8, 0, 0) };
            fileRow.Children.Add(UiKit.Secondary(Strings.T("todo.compose.files.add"), delegate { PickFiles(); }));
            body.Children.Add(fileRow);

            if (attachments != null)
            {
                foreach (var attachment in attachments)
                {
                    if (attachment != null)
                    {
                        _pending.Add(new PendingFile { Record = attachment });
                    }
                }
            }

            RefreshFiles();

            // --- errors and actions ----------------------------------------------
            _error = new TextBlock
            {
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.DangerBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(2, 10, 0, 0),
                Visibility = Visibility.Collapsed
            };
            body.Children.Add(_error);

            var scroll = new ScrollViewer
            {
                Content = body,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Margin = new Thickness(0, 0, 0, 10)
            };

            var actions = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 4, 0, 0)
            };
            var cancel = UiKit.Secondary(Strings.T("todo.compose.cancel"), delegate { _sheet.Close(); });
            cancel.MinWidth = 88;
            cancel.Margin = new Thickness(0, 0, 10, 0);
            var save = UiKit.Primary(Strings.T("todo.compose.save"), delegate { Save(); });
            save.MinWidth = 112;
            actions.Children.Add(cancel);
            actions.Children.Add(save);

            var root = new DockPanel { LastChildFill = true };
            DockPanel.SetDock(actions, Dock.Bottom);
            root.Children.Add(actions);
            root.Children.Add(scroll);

            var title = existing == null ? Strings.T("todo.compose.title") : Strings.T("todo.compose.edit");
            _sheet = UiKit.GlassSheet(title, root, 640, 780);
            _sheet.Owner = owner;
            _sheet.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.Key == Key.Escape)
                {
                    _sheet.Close();
                }
            };
        }

        /// <summary>
        /// Shows the sheet and returns what was entered, or null when it was cancelled.
        /// </summary>
        public static TodoDraft Show(
            Window owner,
            MailSummary mail,
            TodoItem existing,
            IList<TodoAttachment> attachments,
            IList<string> categories,
            string prefillTitle = null)
        {
            var composer = new TodoComposerSheet(owner, mail, existing, attachments, categories, prefillTitle);
            var accepted = composer._sheet.ShowDialog();
            return accepted == true ? composer.Result : null;
        }

        // ------------------------------------------------------------------ pieces

        private static TextBlock SectionLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = UiKit.TypeCaption,
                FontWeight = FontWeights.Medium,
                Foreground = Theme.MutedBrush,
                Margin = new Thickness(2, 12, 0, 5)
            };
        }

        private static Button SmallButton(string text, Action onClick)
        {
            var button = UiKit.Secondary(text, delegate { onClick(); });
            button.Height = 26;
            button.MinWidth = 0;
            button.Padding = new Thickness(10, 0, 10, 0);
            button.Margin = new Thickness(0, 0, 8, 0);
            return button;
        }

        private static DatePicker NewDatePicker(DateTime? utc, string name)
        {
            var picker = new DatePicker
            {
                Height = 34,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Foreground = Theme.InkBrush,
                FontSize = UiKit.TypeBody,
                SelectedDate = TodoDates.DayFromDue(utc),
                VerticalContentAlignment = VerticalAlignment.Center,
                IsTodayHighlighted = true
            };
            AutomationProperties.SetName(picker, name);
            return picker;
        }

        private static string ShortSubject(string subject)
        {
            if (string.IsNullOrEmpty(subject))
            {
                return string.Empty;
            }

            var text = subject.Trim();
            return text.Length <= 40 ? text : text.Substring(0, 40) + "\u2026";
        }

        private string ShortSubject()
        {
            return _mail == null ? string.Empty : ShortSubject(_mail.Subject);
        }

        private static string MailLine(MailSummary mail)
        {
            var from = string.IsNullOrEmpty(mail.FromName) ? mail.FromAddress : mail.FromName;
            if (mail.ReceivedOn == default(DateTime))
            {
                return from ?? string.Empty;
            }

            return string.Format(
                CultureInfo.CurrentCulture,
                Strings.T("todo.compose.mail.line"),
                from,
                mail.ReceivedOn.ToString("yyyy-MM-dd HH:mm"));
        }

        private static DateTime NextMonday()
        {
            var today = DateTime.Now.Date;
            var delta = ((int)DayOfWeek.Monday - (int)today.DayOfWeek + 7) % 7;
            if (delta == 0)
            {
                delta = 7;
            }

            return today.AddDays(delta);
        }

        private void SetDue(DateTime day)
        {
            _due.SelectedDate = day;
            if (_remind.IsChecked == true && !_remindDate.SelectedDate.HasValue)
            {
                _remindDate.SelectedDate = day;
            }
        }

        private void FillNoteFromMail(bool senderOnly)
        {
            if (_mail == null)
            {
                return;
            }

            var from = string.IsNullOrEmpty(_mail.FromName) ? _mail.FromAddress : _mail.FromName;
            var line = string.Format(
                CultureInfo.CurrentCulture,
                Strings.T("todo.compose.mail.line"),
                from,
                _mail.ReceivedOn == default(DateTime) ? string.Empty : _mail.ReceivedOn.ToString("yyyy-MM-dd HH:mm"));

            string text;
            if (senderOnly)
            {
                text = line;
            }
            else
            {
                text = string.IsNullOrEmpty(_mail.Preview)
                    ? (line + Environment.NewLine + (_mail.Subject ?? string.Empty))
                    : (_mail.Preview.Trim() + Environment.NewLine + Environment.NewLine + line);
            }

            // Append rather than overwrite: the user may already have typed the first
            // step, and losing it to a quick-fill button would be a data loss.
            _note.Text = string.IsNullOrEmpty(_note.Text) ? text : (_note.Text.TrimEnd() + Environment.NewLine + Environment.NewLine + text);
        }

        // ------------------------------------------------------------------ files

        private void PickFiles()
        {
            var dialog = new OpenFileDialog
            {
                Multiselect = true,
                Title = Strings.T("todo.compose.files.add"),
                Filter = Strings.T("todo.compose.files.filter") + "|*.*"
            };

            if (dialog.ShowDialog(_sheet) != true)
            {
                return;
            }

            foreach (var path in dialog.FileNames)
            {
                var known = false;
                foreach (var pending in _pending)
                {
                    if (pending.Path != null && string.Equals(pending.Path, path, StringComparison.OrdinalIgnoreCase))
                    {
                        known = true;
                        break;
                    }
                }

                if (!known)
                {
                    _pending.Add(new PendingFile { Path = path });
                }
            }

            RefreshFiles();
        }

        private void RefreshFiles()
        {
            var stack = new StackPanel();

            if (_pending.Count == 0)
            {
                stack.Children.Add(new TextBlock
                {
                    Text = Strings.T("todo.compose.files.none"),
                    FontSize = UiKit.TypeCaption,
                    Foreground = Theme.MutedBrush
                });
            }

            for (var i = 0; i < _pending.Count; i++)
            {
                var index = i;
                var pending = _pending[i];
                var name = pending.Record != null ? pending.Record.FileName : System.IO.Path.GetFileName(pending.Path);
                var size = pending.Record != null ? pending.Record.SizeBytes : SafeLength(pending.Path);

                var row = new DockPanel { LastChildFill = true, Margin = new Thickness(0, 2, 0, 2) };
                var remove = UiKit.Secondary(Strings.T("todo.compose.files.remove"), delegate { RemoveFile(index); });
                remove.Height = 24;
                remove.MinWidth = 0;
                remove.Padding = new Thickness(8, 0, 8, 0);
                remove.FontSize = UiKit.TypeCaption;
                DockPanel.SetDock(remove, Dock.Right);
                row.Children.Add(remove);
                row.Children.Add(new TextBlock
                {
                    Text = string.Format(CultureInfo.CurrentCulture, "{0}  ({1})", name, SizeText(size)),
                    FontSize = UiKit.TypeCaption,
                    Foreground = Theme.InkBrush,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    Margin = new Thickness(0, 0, 8, 0)
                });
                stack.Children.Add(row);
            }

            _fileHost.Child = stack;
        }

        private void RemoveFile(int index)
        {
            if (index < 0 || index >= _pending.Count)
            {
                return;
            }

            var pending = _pending[index];
            if (pending.Record != null && !string.IsNullOrEmpty(pending.Record.StoredName))
            {
                _removed.Add(pending.Record.StoredName);
            }

            _pending.RemoveAt(index);
            RefreshFiles();
        }

        private static long SafeLength(string path)
        {
            try
            {
                return string.IsNullOrEmpty(path) || !File.Exists(path) ? 0 : new FileInfo(path).Length;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static string SizeText(long bytes)
        {
            if (bytes <= 0)
            {
                return "0 KB";
            }

            if (bytes >= 1024L * 1024L)
            {
                return (bytes / (1024.0 * 1024.0)).ToString("0.0", CultureInfo.CurrentCulture) + " MB";
            }

            return Math.Max(1, bytes / 1024L).ToString(CultureInfo.CurrentCulture) + " KB";
        }

        // ------------------------------------------------------------------ save

        private void Save()
        {
            var title = (_title.Text ?? string.Empty).Trim();
            if (title.Length == 0)
            {
                ShowError(Strings.T("todo.compose.titleRequired"));
                _title.Focus();
                return;
            }

            var repeat = _repeat.SelectedIndex > 0 ? (TodoRepeat)_repeat.SelectedIndex : TodoRepeat.None;
            var due = TodoDates.DueFromDay(_due.SelectedDate);
            var reminderDay = _remindDate.SelectedDate ?? _due.SelectedDate;
            var reminder = _remind.IsChecked == true ? TodoDates.DueFromDay(reminderDay) : null;

            var draft = new TodoDraft
            {
                Title = title,
                Note = _note.Text ?? string.Empty,
                Category = (_category.Text ?? string.Empty).Trim(),
                DueOn = due,
                ReminderOn = reminder,
                Repeat = repeat,
                RepeatAnchorDay = (repeat == TodoRepeat.Monthly || repeat == TodoRepeat.Yearly) && due.HasValue
                    ? TodoDates.LocalDayOf(due.Value).Day
                    : 0
            };

            if (_mail != null && _attachMail != null && _attachMail.IsChecked == true)
            {
                draft.SourceEntryId = _mail.EntryId;
                draft.SourceFrom = string.IsNullOrEmpty(_mail.FromName) ? _mail.FromAddress : _mail.FromName;
                draft.SourceReceivedOn = TodoDates.ToUtc(_mail.ReceivedOn);
            }
            else if (_mail == null && _existing != null)
            {
                // No mail in hand (opened from the follow-up list): keep the link the
                // item already had instead of silently orphaning it.
                draft.SourceEntryId = _existing.SourceEntryId;
                draft.SourceFrom = _existing.SourceFrom;
                draft.SourceReceivedOn = _existing.SourceReceivedOn;
            }

            foreach (var pending in _pending)
            {
                if (pending.Record == null && !string.IsNullOrEmpty(pending.Path))
                {
                    draft.NewFiles.Add(pending.Path);
                }
            }

            foreach (var stored in _removed)
            {
                draft.RemovedStoredNames.Add(stored);
            }

            Result = draft;
            _sheet.DialogResult = true;
        }

        private void ShowError(string text)
        {
            _error.Text = text;
            _error.Visibility = Visibility.Visible;
        }
    }
}

