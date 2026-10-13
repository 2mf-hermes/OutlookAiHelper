using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// The detail panel, grouped into cards instead of one long column of loose rows.
    ///
    /// Every card answers one question: what the mail is, why it was filed where it was, what else
    /// was said in its topic, what has been recorded about it, and what can be done with it. Fields
    /// that belong together sit in one card, and a fact is written once -- the panel used to print
    /// the same arrival time and due date in two different shapes.
    ///
    /// The chrome is the settings page's own vocabulary (UiKit.InsetGroup: a small caption over a
    /// white rounded card), so the panel and the settings page stay the same material.
    ///
    /// Kept separate from MainWindow.cs, which is long already; it is a partial because it reads
    /// the panel's fields and a caller's selection out of the rest of the window.
    /// </summary>
    public sealed partial class MainWindow
    {
        // ------------------------------------------------------------------ card pieces

        /// <summary>
        /// One card of the panel: a caption over a white rounded list. Returns the block itself, so
        /// the caller can take the caption down with the card when there is nothing to show.
        /// </summary>
        private static StackPanel DetailBlock(string header, UIElement rows)
        {
            var block = (StackPanel)UiKit.InsetGroup(header, rows);
            block.Margin = new Thickness(0, 0, 0, 14);
            return block;
        }

        /// <summary>
        /// The hairline between two rows of a card. UiKit.AddSeparated draws its line in white,
        /// which is invisible on a white card, so rows here are ruled in the list separator colour.
        /// </summary>
        private static Border DetailRule()
        {
            return new Border
            {
                Height = 1,
                Margin = new Thickness(16, 0, 16, 0),
                Background = UiKit.SeparatorBrush()
            };
        }

        private static void AddDetailRow(StackPanel rows, FrameworkElement row, bool isLast)
        {
            rows.Children.Add(row);
            if (!isLast)
            {
                rows.Children.Add(DetailRule());
            }
        }

        /// <summary>
        /// One fact: its name in small text, its value underneath. This is how a system inspector
        /// labels a field, and it stays readable when the value is a folder path.
        /// </summary>
        private static FrameworkElement DetailFact(string label, UIElement value)
        {
            var row = new StackPanel { Margin = new Thickness(16, 10, 16, 12) };
            row.Children.Add(new TextBlock
            {
                Text = label,
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.NoWrap
            });

            var text = value as TextBlock;
            if (text != null)
            {
                text.Margin = new Thickness(0, 2, 0, 0);
            }

            row.Children.Add(value);
            return row;
        }

        private static TextBlock DetailValue(string text)
        {
            return DetailValue(text, Theme.InkBrush);
        }

        private static TextBlock DetailValue(string text, Brush ink)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = UiKit.TypeSubhead,
                Foreground = ink,
                TextWrapping = TextWrapping.Wrap
            };
        }

        /// <summary>A state worth pulling out of the text: a tinted capsule.</summary>
        private static Border DetailChip(string text, Brush ink)
        {
            return UiKit.Badge(text, Theme.BadgeTintBrush, ink);
        }

        // ------------------------------------------------------------------ the panel's cards

        /// <summary>The card of facts about the mail: who sent it, when, where it lives, what it is.</summary>
        private StackPanel BuildMailFactsBlock()
        {
            _mailFactsRows = UiKit.ListRows();
            return DetailBlock(Strings.T("detail.section.mail"), _mailFactsRows);
        }

        /// <summary>
        /// The card that says why the mail was filed where it was. Its rows are written per
        /// selection by the renderer, and the whole card leaves the page when there is no reason
        /// to show.
        /// </summary>
        private StackPanel BuildReasonBlock()
        {
            _reasonPanel = UiKit.ListRows();
            return DetailBlock(Strings.T("reason.title"), _reasonPanel);
        }

        /// <summary>
        /// What can be done with what the panel is showing, in one card, so the buttons stop
        /// competing with the mail's own facts for the bottom of the column.
        /// </summary>
        private StackPanel BuildActionBlock()
        {
            var buttons = new StackPanel { Margin = new Thickness(16, 4, 16, 14) };
            buttons.Children.Add(_reclassifyButton);
            buttons.Children.Add(_addTodoButton);
            buttons.Children.Add(_aiAction);
            return DetailBlock(Strings.T("detail.section.actions"), buttons);
        }

        // ------------------------------------------------------------------ the mail card

        /// <summary>Fills the mail card from the selected mail, or takes the card down if there is none.</summary>
        private void FillMailFacts(ScanResultItem item)
        {
            if (_mailFactsRows == null || _mailFactsBlock == null)
            {
                return;
            }

            _mailFactsRows.Children.Clear();
            if (item == null || item.Mail == null)
            {
                _mailFactsBlock.Visibility = Visibility.Collapsed;
                return;
            }

            var rows = new List<FrameworkElement>(MailFactRows(item));
            for (var i = 0; i < rows.Count; i++)
            {
                AddDetailRow(_mailFactsRows, rows[i], i == rows.Count - 1);
            }

            _mailFactsBlock.Visibility = Visibility.Visible;
        }

        /// <summary>Empties the mail card and takes it down; the follow-up-only view has no mail to describe.</summary>
        private void ClearMailFacts()
        {
            if (_mailFactsRows != null)
            {
                _mailFactsRows.Children.Clear();
            }

            if (_mailFactsBlock != null)
            {
                _mailFactsBlock.Visibility = Visibility.Collapsed;
            }
        }

        /// <summary>The facts of the mail itself, in the order the panel reads them.</summary>
        private static IEnumerable<FrameworkElement> MailFactRows(ScanResultItem item)
        {
            var mail = item.Mail;

            yield return DetailFact(Strings.T("detail.from"), DetailValue(mail.FromName));
            yield return DetailFact(Strings.T("detail.time"), DetailValue(mail.ReceivedOn.ToString("yyyy-MM-dd HH:mm")));
            yield return DetailFact(Strings.T("detail.folder"), DetailValue(mail.FolderPath));
            yield return DetailFact(
                Strings.T("todo.detail.quadrant"),
                DetailValue(
                    Strings.QuadrantName(item.Classification.Quadrant),
                    Theme.QuadrantBrush(item.Classification.Quadrant)));
            yield return DetailFact(Strings.T("detail.state"), MailStateRow(item));
        }

        /// <summary>
        /// A mail's state as capsules: one per thing that is true of it. A mail with nothing to
        /// report gets a single quiet capsule instead of an empty row, so the field never looks
        /// like it failed to load.
        /// </summary>
        private static FrameworkElement MailStateRow(ScanResultItem item)
        {
            var mail = item.Mail;
            var chips = new WrapPanel { Margin = new Thickness(0, 4, 0, 0) };

            AddMailStateChip(chips, !mail.IsRead, Strings.T("reason.signal.unread"));
            AddMailStateChip(chips, mail.HasFlag, Strings.T("reason.signal.flag"));
            AddMailStateChip(chips, mail.IsMeetingRequest, Strings.T("reason.signal.meeting"));
            AddMailStateChip(chips, mail.Importance >= 2, Strings.T("reason.signal.importance"));

            if (chips.Children.Count == 0)
            {
                AddMailStateChip(chips, true, Strings.T("reason.signal.default"));
            }

            return chips;
        }

        private static void AddMailStateChip(WrapPanel chips, bool when, string text)
        {
            if (!when)
            {
                return;
            }

            var chip = DetailChip(text, Theme.AccentBrush);
            chip.Margin = new Thickness(0, 0, 6, 6);
            chips.Children.Add(chip);
        }

        // ------------------------------------------------------------------ the follow-up card

        /// <summary>Fills the follow-up card from the open item; nothing is added while a plain mail is shown.</summary>
        private void FillTodoFacts(TodoItem item)
        {
            if (_todoFactsRows == null)
            {
                return;
            }

            _todoFactsRows.Children.Clear();
            if (item == null)
            {
                return;
            }

            var rows = new List<FrameworkElement>(TodoFactRows(item));
            for (var i = 0; i < rows.Count; i++)
            {
                AddDetailRow(_todoFactsRows, rows[i], i == rows.Count - 1);
            }
        }

        /// <summary>What has been recorded about the follow-up, as opposed to what the mail says.</summary>
        private static IEnumerable<FrameworkElement> TodoFactRows(TodoItem item)
        {
            var done = item.Status == TodoStatus.Done;
            yield return DetailFact(
                Strings.T("todo.detail.state"),
                DetailChip(
                    Strings.T(done ? "todo.filter.done" : "todo.filter.open"),
                    done ? Theme.MutedBrush : Theme.AccentBrush));

            if (!string.IsNullOrEmpty(item.QuadrantHint))
            {
                yield return DetailFact(Strings.T("todo.detail.quadrant"), DetailValue(item.QuadrantHint));
            }

            yield return DetailFact(
                Strings.T("todo.detail.created"),
                DetailValue(item.CreatedOn.ToLocalTime().ToString("yyyy-MM-dd HH:mm")));
            yield return DetailFact(
                Strings.T("todo.detail.updated"),
                DetailValue(item.UpdatedOn.ToLocalTime().ToString("yyyy-MM-dd HH:mm")));

            if (item.DueOn.HasValue)
            {
                yield return DetailFact(Strings.T("todo.detail.due"), DetailValue(DescribeDue(item)));
            }
        }

        // ------------------------------------------------------------------ the reason card

        /// <summary>Shows or hides the reason card; an empty card reads as a bug, so it leaves the page.</summary>
        private void SetReasonBlockVisible(bool visible)
        {
            if (_reasonBlock != null)
            {
                _reasonBlock.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
            }
        }
    }
}
