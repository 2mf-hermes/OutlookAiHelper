using System;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// Search for the two lists. Both fields narrow what is already in hand - the scan result and
    /// the follow-up list - so a search costs no Outlook call, keeps working with Outlook closed,
    /// and cannot reach a mail the scan did not bring back. A query lives only in its text box:
    /// it is never stored, logged or sent.
    /// </summary>
    public partial class MainWindow
    {
        private TextBox _mailSearch;
        private TextBox _todoSearch;
        private TextBlock _listSearchEmpty;

        /// <summary>
        /// The mail list's search field, inside the list's own card so it reads as part of the
        /// list it narrows rather than as another toolbar command.
        /// </summary>
        private UIElement BuildMailSearchRow()
        {
            TextBox box;
            var field = BuildSearchField("search.mail.placeholder", out box, OnMailSearchChanged);
            _mailSearch = box;

            var row = new StackPanel { Margin = new Thickness(12, 10, 12, 6) };
            row.Children.Add(field);
            return row;
        }

        /// <summary>
        /// The follow-up page's search field, under the controls that decide which follow-ups are
        /// on the page and closest to the list it narrows.
        /// </summary>
        private UIElement BuildTodoSearchRow()
        {
            TextBox box;
            var field = BuildSearchField("search.todo.placeholder", out box, OnTodoSearchChanged);
            _todoSearch = box;

            var row = new StackPanel { Margin = new Thickness(0, 8, 0, 0) };
            row.Children.Add(field);
            return row;
        }

        /// <summary>
        /// One search field: the same glass input the rest of the app uses, a hint standing in for
        /// a placeholder (WPF has none) that steps aside as soon as there is something to read, and
        /// Escape to clear - what a search field is expected to do everywhere else. The hint is
        /// also the field's accessible name, so a screen reader announces what can be searched
        /// instead of a bare "edit".
        /// </summary>
        private static FrameworkElement BuildSearchField(string hintKey, out TextBox box, Action changed)
        {
            var field = UiKit.Input(string.Empty);
            field.Height = 34;
            field.ToolTip = Strings.T("search.help");
            AutomationProperties.SetName(field, Strings.T(hintKey));
            box = field;

            var hint = new TextBlock
            {
                Text = Strings.T(hintKey),
                FontSize = UiKit.TypeFootnote,
                Foreground = Theme.MutedBrush,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(14, 0, 14, 0),
                IsHitTestVisible = false,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            // Escape is only swallowed when it actually emptied the field, so it keeps its usual
            // meaning in a window that has nothing to clear.
            field.TextChanged += (s, e) =>
            {
                hint.Visibility = field.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
                if (changed != null)
                {
                    changed();
                }
            };
            field.KeyDown += (s, e) =>
            {
                if (e.Key == Key.Escape && field.Text.Length > 0)
                {
                    field.Text = string.Empty;
                    e.Handled = true;
                }
            };

            var host = new Grid { Height = 34, MaxWidth = 420, HorizontalAlignment = HorizontalAlignment.Left };
            host.Children.Add(field);
            host.Children.Add(hint);
            return host;
        }

        private static string QueryOf(TextBox box)
        {
            return box == null ? null : box.Text;
        }

        /// <summary>
        /// Whether a mail answers the mail list's query: subject, the opening of its text, both
        /// sender spellings and either recipient line. On the cloud path only the subject and the
        /// sender exist - the reader asks for Mail.ReadBasic, which carries no body and no
        /// recipients - so those searches simply match nothing there rather than pretending.
        /// </summary>
        private bool MatchesMail(MailSummary mail)
        {
            if (mail == null)
            {
                return false;
            }

            return ListSearch.Matches(
                QueryOf(_mailSearch),
                mail.Subject,
                mail.Preview,
                mail.FromName,
                mail.FromAddress,
                mail.ToLine,
                mail.CcLine);
        }

        /// <summary>
        /// Whether a follow-up answers the follow-up page's query. A follow-up carries a title, a
        /// note and the sender it was made from, so the mail it came from is searched as well -
        /// that is where a recipient or a body word lives, whenever that mail is still scanned.
        /// </summary>
        private bool MatchesTodo(TodoItem item)
        {
            if (item == null)
            {
                return false;
            }

            var query = QueryOf(_todoSearch);
            if (!ListSearch.IsActive(query))
            {
                return true;
            }

            if (ListSearch.Matches(query, item.Title, item.Note, item.SourceFrom))
            {
                return true;
            }

            var source = FindScannedMail(item.SourceEntryId);
            return source != null
                && source.Mail != null
                && ListSearch.Matches(
                    query,
                    source.Mail.Subject,
                    source.Mail.Preview,
                    source.Mail.FromName,
                    source.Mail.FromAddress,
                    source.Mail.ToLine,
                    source.Mail.CcLine);
        }

        /// <summary>
        /// A keystroke changes what the list shows, never what was scanned: the same rows come
        /// back the moment the field is cleared, with no trip to Outlook.
        /// </summary>
        private void OnMailSearchChanged()
        {
            if (_mailList == null)
            {
                return;
            }

            BindList();
        }

        private void OnTodoSearchChanged()
        {
            if (_todoList == null)
            {
                return;
            }

            BindTodos();
        }
    }
}
