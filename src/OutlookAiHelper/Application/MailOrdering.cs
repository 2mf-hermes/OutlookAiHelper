using System;
using System.Collections.Generic;
using System.Linq;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Application
{
    /// <summary>
    /// How the mail list — and the mails shown beside a follow-up — are ordered.
    ///
    /// NewestFirst is the default and the reference point: a folded topic row is read as
    /// "when did something last happen in this topic", so a row's date is its newest mail
    /// and not the one it presents. Every other mode keeps newest-first for ties, so two
    /// rows that share a sender or a quadrant still read top-down by time.
    /// </summary>
    public enum MailSortMode
    {
        NewestFirst,
        OldestFirst,
        Quadrant,
        Sender,
        Subject
    }

    /// <summary>
    /// The one ordering rule behind both the mail list and the follow-up panel's related
    /// mails. Those two surfaces share a single setting, so they have to share the comparer
    /// as well — otherwise the same mail could sit in a different place on each.
    ///
    /// Kept out of the window so the order is testable without a mailbox: the unit tests
    /// hand it lists of mails and read back the order.
    /// </summary>
    public static class MailOrdering
    {
        public const MailSortMode DefaultMode = MailSortMode.NewestFirst;

        /// <summary>
        /// Modes in the order the settings picker offers them. This array is the picker's
        /// item source, so another mode only has to be added here and given a label.
        /// </summary>
        public static readonly MailSortMode[] Modes =
        {
            MailSortMode.NewestFirst,
            MailSortMode.OldestFirst,
            MailSortMode.Quadrant,
            MailSortMode.Sender,
            MailSortMode.Subject
        };

        /// <summary>
        /// Mode for a stored name. Anything unrecognised — an older settings file, a
        /// hand-edited value — reads as the default instead of failing, the same way an
        /// unknown follow-up sort does.
        /// </summary>
        public static MailSortMode ParseMode(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return DefaultMode;
            }

            foreach (var mode in Modes)
            {
                if (string.Equals(mode.ToString(), value.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return mode;
                }
            }

            return DefaultMode;
        }

        /// <summary>Label key for a mode, so both surfaces name it the same way.</summary>
        public static string LabelKey(MailSortMode mode)
        {
            switch (mode)
            {
                case MailSortMode.OldestFirst:
                    return "mail.sort.oldest";
                case MailSortMode.Quadrant:
                    return "mail.sort.quadrant";
                case MailSortMode.Sender:
                    return "mail.sort.sender";
                case MailSortMode.Subject:
                    return "mail.sort.subject";
                default:
                    return "mail.sort.newest";
            }
        }

        /// <summary>
        /// Mails in display order. A null item, or one with no summary behind it, is dropped
        /// rather than sorted: it has no date, no sender and no quadrant to be placed by.
        /// </summary>
        public static List<ScanResultItem> SortMails(IEnumerable<ScanResultItem> items, MailSortMode mode)
        {
            var list = (items ?? Enumerable.Empty<ScanResultItem>())
                .Where(i => i != null && i.Mail != null)
                .ToList();
            if (list.Count <= 1)
            {
                return list;
            }

            list.Sort(ComparerFor(mode));
            return list;
        }

        /// <summary>
        /// Topic rows in display order. The time modes read the topic's newest mail, because
        /// that is the date the row carries; every other mode reads the mail the row
        /// presents, so a topic keeps the sender and quadrant the user sees on it.
        /// </summary>
        public static List<MailListRow> SortRows(IEnumerable<MailListRow> rows, MailSortMode mode)
        {
            var list = (rows ?? Enumerable.Empty<MailListRow>())
                .Where(r => r != null)
                .ToList();
            if (list.Count <= 1)
            {
                return list;
            }

            var comparer = ComparerFor(mode);
            var byNewest = mode == MailSortMode.NewestFirst || mode == MailSortMode.OldestFirst;
            return byNewest
                ? list.OrderBy(r => r.Newest ?? r.Head, comparer).ToList()
                : list.OrderBy(r => r.Head ?? r.Newest, comparer).ToList();
        }

        public static IComparer<ScanResultItem> ComparerFor(MailSortMode mode)
        {
            return new MailComparer(mode);
        }

        public static DateTime ReceivedOn(MailSummary mail)
        {
            return mail == null ? DateTime.MinValue : mail.ReceivedOn;
        }

        /// <summary>0 = urgent and important; the enum order is the rank order.</summary>
        public static int QuadrantRank(ScanResultItem item)
        {
            if (item == null || item.Classification == null)
            {
                return 9;
            }

            return (int)item.Classification.Quadrant;
        }

        /// <summary>Sender as one sortable string: the display name, then the address.</summary>
        public static string SenderKey(MailSummary mail)
        {
            if (mail == null)
            {
                return string.Empty;
            }

            var name = mail.FromName ?? string.Empty;
            var address = mail.FromAddress ?? string.Empty;
            return name.Length > 0 ? name + " " + address : address;
        }

        /// <summary>Topic identity without the reply markers, so a thread groups under one key.</summary>
        public static string SubjectKey(MailSummary mail)
        {
            return mail == null ? string.Empty : MailThreads.StripSubject(mail.Subject);
        }

        /// <summary>
        /// Comparer behind both surfaces. Ties fall through to the subject and then to the
        /// newest arrival, which makes the order total: two binds of the same list cannot
        /// swap rows around, so the list does not flicker when it is redrawn.
        /// </summary>
        private sealed class MailComparer : IComparer<ScanResultItem>
        {
            private readonly MailSortMode _mode;

            public MailComparer(MailSortMode mode)
            {
                _mode = mode;
            }

            public int Compare(ScanResultItem x, ScanResultItem y)
            {
                if (ReferenceEquals(x, y))
                {
                    return 0;
                }

                // A topic row with nothing behind it sorts to the end instead of throwing.
                if (x == null)
                {
                    return 1;
                }

                if (y == null)
                {
                    return -1;
                }

                var primary = ComparePrimary(x, y);
                if (primary != 0)
                {
                    return primary;
                }

                if (_mode == MailSortMode.Subject)
                {
                    return CompareNewestFirst(x, y);
                }

                var bySubject = string.Compare(
                    SubjectKey(x.Mail),
                    SubjectKey(y.Mail),
                    StringComparison.CurrentCultureIgnoreCase);
                return bySubject != 0 ? bySubject : CompareNewestFirst(x, y);
            }

            private int ComparePrimary(ScanResultItem x, ScanResultItem y)
            {
                switch (_mode)
                {
                    case MailSortMode.OldestFirst:
                        return ReceivedOn(x.Mail).CompareTo(ReceivedOn(y.Mail));
                    case MailSortMode.Quadrant:
                        return QuadrantRank(x).CompareTo(QuadrantRank(y));
                    case MailSortMode.Sender:
                        return string.Compare(
                            SenderKey(x.Mail),
                            SenderKey(y.Mail),
                            StringComparison.CurrentCultureIgnoreCase);
                    case MailSortMode.Subject:
                        return string.Compare(
                            SubjectKey(x.Mail),
                            SubjectKey(y.Mail),
                            StringComparison.CurrentCultureIgnoreCase);
                    default:
                        return CompareNewestFirst(x, y);
                }
            }

            private static int CompareNewestFirst(ScanResultItem x, ScanResultItem y)
            {
                return ReceivedOn(y.Mail).CompareTo(ReceivedOn(x.Mail));
            }
        }
    }
}
