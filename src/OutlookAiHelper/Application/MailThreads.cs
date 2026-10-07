using System;
using System.Collections.Generic;
using System.Text;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Application
{
    /// <summary>
    /// One line of the mail list: either a single mail, or the aggregation of every mail that
    /// belongs to the same topic. The window renders one of these per row and never decides
    /// the grouping itself.
    /// </summary>
    public sealed class MailListRow
    {
        public MailListRow()
        {
            Mails = new List<ScanResultItem>();
        }

        /// <summary>Topic identity shared by the members; null when the row is not a thread.</summary>
        public string ThreadKey { get; set; }

        /// <summary>Members in the caller's order (highest-ranked first).</summary>
        public IList<ScanResultItem> Mails { get; private set; }

        /// <summary>Highest-ranked member — the mail the row presents.</summary>
        public ScanResultItem Head { get; set; }

        /// <summary>Most recently received member, i.e. the newest reply in the topic.</summary>
        public ScanResultItem Newest { get; set; }

        /// <summary>How many members were added to the follow-up list.</summary>
        public int TodoCount { get; set; }

        /// <summary>True only for a row that actually folds several mails together.</summary>
        public bool IsThread
        {
            get { return ThreadKey != null && Mails.Count > 1; }
        }

        public bool Expanded { get; set; }

        public bool HasTodo
        {
            get { return TodoCount > 0; }
        }
    }

    /// <summary>Counts shown above the list. They describe the rows currently on screen.</summary>
    public sealed class MailListStats
    {
        public int Mails { get; set; }
        public int Threads { get; set; }
        public int WithTodo { get; set; }
    }

    /// <summary>
    /// Folds scanned mails into topic rows: mails that share a subject (reply and forward
    /// prefixes stripped) are one row, and mail with no usable subject falls back to its
    /// sender so a nameless conversation still reads as one row. Nothing here touches Outlook
    /// or WPF, so the rule is unit-testable on its own.
    /// </summary>
    public static class MailThreads
    {
        /// <summary>
        /// Reply/forward markers, lower-case, with both the ASCII and the full-width colon.
        /// Only these are stripped: anything else is part of the subject.
        /// </summary>
        private static readonly string[] ReplyPrefixes =
        {
            "re:", "fw:", "fwd:", "tr:",
            "\u56de\u8986:", "\u56de\u8986\uff1a", "\u56de\u5fa9:", "\u56de\u5fa9\uff1a",
            "\u7b54\u8986:", "\u7b54\u8986\uff1a", "\u7b54\u590d:", "\u7b54\u590d\uff1a",
            "\u8f49\u5bc4:", "\u8f49\u5bc4\uff1a", "\u8f6c\u53d1:", "\u8f6c\u53d1\uff1a",
            "\u8f49\u767c:", "\u8f49\u767c\uff1a"
        };

        /// <summary>
        /// Strips every leading reply/forward marker and collapses whitespace, keeping the
        /// original casing so the result can be shown as the topic's title.
        /// </summary>
        public static string StripSubject(string subject)
        {
            if (string.IsNullOrEmpty(subject))
            {
                return string.Empty;
            }

            var text = Collapse(subject.Trim());
            var stripped = true;
            while (stripped)
            {
                stripped = false;
                foreach (var prefix in ReplyPrefixes)
                {
                    if (text.Length >= prefix.Length && text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    {
                        text = Collapse(text.Substring(prefix.Length));
                        stripped = true;
                        break;
                    }
                }
            }

            return text;
        }

        /// <summary>Topic identity: the stripped subject, case- and whitespace-insensitive.</summary>
        public static string NormalizeSubject(string subject)
        {
            return StripSubject(subject).ToLowerInvariant();
        }

        /// <summary>
        /// The row key for a mail: the shared subject when there is one, otherwise the sender.
        /// Null means the mail cannot be recognised as part of anything, so it always stands
        /// alone. A one-character subject is treated as unusable — it groups far more often
        /// by accident than on purpose.
        /// </summary>
        public static string KeyOf(MailSummary mail)
        {
            if (mail == null)
            {
                return null;
            }

            var subject = NormalizeSubject(mail.Subject);
            if (subject.Length >= 2)
            {
                return "s:" + subject;
            }

            var sender = SenderKey(mail);
            return sender.Length == 0 ? null : "f:" + sender;
        }

        /// <summary>
        /// Folds an already-ordered list into rows. The caller's order decides both where a
        /// topic appears (its highest-ranked member wins) and the order inside it.
        /// <paramref name="hasTodo"/> answers "was this mail added to the follow-up list?";
        /// <paramref name="isExpanded"/> answers "is this topic open?".
        /// </summary>
        public static IList<MailListRow> Group(
            IEnumerable<ScanResultItem> ordered,
            Func<string, bool> hasTodo,
            Func<string, bool> isExpanded)
        {
            var rows = new List<MailListRow>();
            var index = new Dictionary<string, MailListRow>(StringComparer.Ordinal);

            if (ordered == null)
            {
                return rows;
            }

            foreach (var item in ordered)
            {
                if (item == null)
                {
                    continue;
                }

                var key = KeyOf(item.Mail);
                MailListRow row = null;
                if (key != null)
                {
                    index.TryGetValue(key, out row);
                }

                if (row == null)
                {
                    row = new MailListRow
                    {
                        ThreadKey = key,
                        Head = item,
                        Newest = item,
                        Expanded = key != null && isExpanded != null && isExpanded(key),
                        TodoCount = MailHasTodo(item, hasTodo) ? 1 : 0
                    };
                    row.Mails.Add(item);
                    if (key != null)
                    {
                        index[key] = row;
                    }

                    rows.Add(row);
                    continue;
                }

                row.Mails.Add(item);
                if (MailHasTodo(item, hasTodo))
                {
                    row.TodoCount = row.TodoCount + 1;
                }

                if (ReceivedOn(item) > ReceivedOn(row.Newest))
                {
                    row.Newest = item;
                }
            }

            return rows;
        }

        /// <summary>
        /// Entry ids that are already on the follow-up list, so the mail list can mark them.
        /// This lives here rather than in the window so "which mails carry the 已加入待辦 tag"
        /// is testable without WPF or Outlook.
        /// </summary>
        public static HashSet<string> FollowUpEntryIds(IEnumerable<TodoItem> todos)
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            if (todos == null)
            {
                return ids;
            }

            foreach (var todo in todos)
            {
                if (todo != null && !string.IsNullOrEmpty(todo.SourceEntryId))
                {
                    ids.Add(todo.SourceEntryId);
                }
            }

            return ids;
        }

        /// <summary>Counts the rows as shown, so the header always matches the list.</summary>
        public static MailListStats Summarize(IEnumerable<MailListRow> rows)
        {
            var stats = new MailListStats();
            if (rows == null)
            {
                return stats;
            }

            foreach (var row in rows)
            {
                if (row == null)
                {
                    continue;
                }

                stats.Mails += row.Mails.Count;
                stats.WithTodo += row.TodoCount;
                if (row.IsThread)
                {
                    stats.Threads = stats.Threads + 1;
                }
            }

            return stats;
        }

        /// <summary>
        /// Collapses whitespace runs to a single space. Outlook subjects arrive with tabs and
        /// doubled spaces, which would otherwise split one topic across several rows.
        /// </summary>
        public static string Collapse(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var sb = new StringBuilder(text.Length);
            var pendingSpace = false;
            foreach (var ch in text)
            {
                if (char.IsWhiteSpace(ch))
                {
                    pendingSpace = sb.Length > 0;
                    continue;
                }

                if (pendingSpace)
                {
                    sb.Append(' ');
                    pendingSpace = false;
                }

                sb.Append(ch);
            }

            return sb.ToString();
        }

        private static DateTime ReceivedOn(ScanResultItem item)
        {
            return item == null || item.Mail == null ? DateTime.MinValue : item.Mail.ReceivedOn;
        }

        private static bool MailHasTodo(ScanResultItem item, Func<string, bool> hasTodo)
        {
            if (hasTodo == null || item == null || item.Mail == null || string.IsNullOrEmpty(item.Mail.EntryId))
            {
                return false;
            }

            return hasTodo(item.Mail.EntryId);
        }

        /// <summary>
        /// The sender as a grouping key. The address is preferred: a display name that reads
        /// "Amy" today can read "Amy Chen" tomorrow, and the address is what actually identifies
        /// the sender.
        /// </summary>
        private static string SenderKey(MailSummary mail)
        {
            var sender = mail.FromAddress;
            if (string.IsNullOrEmpty(sender))
            {
                sender = mail.FromName;
            }

            return string.IsNullOrEmpty(sender) ? string.Empty : Collapse(sender.Trim()).ToLowerInvariant();
        }
    }
}
