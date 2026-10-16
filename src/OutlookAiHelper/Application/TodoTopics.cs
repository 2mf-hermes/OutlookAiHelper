using System;
using System.Collections.Generic;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Application
{
    /// <summary>
    /// One row of the follow-up list: either a single follow-up, or a topic — the follow-ups
    /// about one subject — presented together. The window renders one of these per row and
    /// never decides the folding itself.
    /// </summary>
    public sealed class TodoTopicRow
    {
        public TodoTopicRow()
        {
            Items = new List<TodoItem>();
        }

        /// <summary>Topic identity shared by the members; null when the row is not a topic.</summary>
        public string TopicKey { get; set; }

        /// <summary>Members in the caller's order (highest-ranked first).</summary>
        public IList<TodoItem> Items { get; private set; }

        /// <summary>Highest-ranked member — the follow-up the row presents and opens.</summary>
        public TodoItem Lead { get; set; }

        /// <summary>Most recently arrived member, so a topic can still say when it moved last.</summary>
        public TodoItem Newest { get; set; }

        /// <summary>
        /// Member that has been waiting longest. For a topic this is what "waiting" means: the
        /// oldest request in it is the one the user owes the oldest answer to.
        /// </summary>
        public TodoItem Oldest { get; set; }

        /// <summary>How many members are still open.</summary>
        public int OpenCount { get; set; }

        /// <summary>
        /// A row that grouped more than one follow-up. A one-member row is a single follow-up:
        /// there is nothing in it to fold and nothing to unfold, so it is shown as a plain row
        /// and the tally does not count it as a grouping either.
        /// </summary>
        public bool IsTopic
        {
            get { return TopicKey != null && Items.Count > 1; }
        }

        /// <summary>Is this topic showing its members?</summary>
        public bool Expanded { get; set; }
    }

    /// <summary>Counts the follow-up list as shown, so the tally above it always matches.</summary>
    public sealed class TodoTopicStats
    {
        /// <summary>Follow-ups listed, counted once each however they are folded.</summary>
        public int Items { get; set; }

        /// <summary>Topics that really grouped several follow-ups.</summary>
        public int Topics { get; set; }

        /// <summary>How many of the listed follow-ups are finished.</summary>
        public int Done { get; set; }
    }

    /// <summary>
    /// Folds follow-ups about the same subject into one row. The topic is the follow-up's title
    /// with reply and forward prefixes stripped, keyed by the very same rule the mail list folds
    /// mail with — so a topic means the same thing on both pages — and a follow-up whose subject
    /// cannot carry a topic stands on its own. A follow-up from a mail carries that mail's
    /// subject as its title, so the two agree; nothing here touches WPF, so the rule is
    /// unit-testable on its own.
    /// </summary>
    public static class TodoTopics
    {
        /// <summary>
        /// The key a follow-up folds under, or null when its title cannot carry a topic. Done
        /// and open follow-ups of one subject share the key on purpose: a topic that still asks
        /// for an answer should show that it is partly answered, rather than hiding the one
        /// finished follow-up the user would have recognised it by.
        /// </summary>
        public static string KeyOf(TodoItem item)
        {
            return item == null ? null : MailThreads.KeyOfSubject(item.Title);
        }

        /// <summary>
        /// Folds an already-ordered list into rows. The caller's order decides both where a
        /// topic appears (its highest-ranked member wins) and the order inside it, so the sort
        /// the user picked still governs the whole list.
        /// </summary>
        /// <param name="ordered">Follow-ups already in the order the page is to show them.</param>
        /// <param name="isExpanded">Answers "is this topic open?" for a topic key.</param>
        public static IList<TodoTopicRow> Group(IEnumerable<TodoItem> ordered, Func<string, bool> isExpanded)
        {
            var rows = new List<TodoTopicRow>();
            var byKey = new Dictionary<string, TodoTopicRow>(StringComparer.Ordinal);

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

                var key = KeyOf(item);
                TodoTopicRow row = null;
                if (key != null)
                {
                    byKey.TryGetValue(key, out row);
                }

                if (row == null)
                {
                    row = new TodoTopicRow
                    {
                        TopicKey = key,
                        Lead = item,
                        Newest = item,
                        Oldest = item,
                        OpenCount = item.Status == TodoStatus.Open ? 1 : 0,
                        Expanded = key != null && isExpanded != null && isExpanded(key)
                    };
                    row.Items.Add(item);
                    if (key != null)
                    {
                        byKey[key] = row;
                    }

                    rows.Add(row);
                    continue;
                }

                row.Items.Add(item);
                if (item.Status == TodoStatus.Open)
                {
                    row.OpenCount = row.OpenCount + 1;
                }

                if (TodoOrdering.WaitingSince(item) > TodoOrdering.WaitingSince(row.Newest))
                {
                    row.Newest = item;
                }

                if (TodoOrdering.WaitingSince(item) < TodoOrdering.WaitingSince(row.Oldest))
                {
                    row.Oldest = item;
                }
            }

            return rows;
        }

        /// <summary>
        /// Counts the rows as shown: the follow-ups listed, the topics they folded into, and how
        /// many of the listed follow-ups are finished. Each member is counted once however the
        /// fold came out, so the tally cannot disagree with the rows below it.
        /// </summary>
        public static TodoTopicStats Summarize(IEnumerable<TodoTopicRow> rows)
        {
            var stats = new TodoTopicStats();
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

                stats.Items += row.Items.Count;
                stats.Done += row.Items.Count - row.OpenCount;
                if (row.IsTopic)
                {
                    stats.Topics = stats.Topics + 1;
                }
            }

            return stats;
        }
    }
}
