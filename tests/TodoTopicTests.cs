using System;
using System.Collections.Generic;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// The follow-up list's topic fold, measured on its own. The rule is the mail list's own
    /// subject rule, so it can be read and checked without a window, Outlook or a store, and the
    /// two pages cannot drift apart without one of these failing.
    /// </summary>
    public static class TodoTopicTests
    {
        public static int Run()
        {
            var failed = 0;
            var t0 = new DateTime(2026, 10, 1, 12, 0, 0, DateTimeKind.Utc);

            failed += Check("follow-ups about one subject fold into a single row", () =>
            {
                var rows = TodoTopics.Group(new List<TodoItem>
                {
                    Item("Report", "e1", t0.AddDays(-2)),
                    Item("RE: Report", "e2", t0),
                    Item("RE: RE: Report", "e3", t0.AddDays(-1))
                }, null);

                return rows.Count == 1 && rows[0].IsTopic && rows[0].Items.Count == 3;
            });

            failed += Check("follow-ups about different subjects stay apart", () =>
            {
                var rows = TodoTopics.Group(new List<TodoItem>
                {
                    Item("Report", "e1", t0),
                    Item("Invoice", "e2", t0)
                }, null);

                return rows.Count == 2
                    && !rows[0].IsTopic && !rows[1].IsTopic
                    && rows[0].Items.Count == 1 && rows[1].Items.Count == 1;
            });

            failed += Check("case and stray spacing do not split a topic", () =>
            {
                var rows = TodoTopics.Group(new List<TodoItem>
                {
                    Item("Report", "e1", t0),
                    Item("  report  ", "e2", t0)
                }, null);

                return rows.Count == 1 && rows[0].IsTopic && rows[0].Items.Count == 2;
            });

            failed += Check("a subject nobody replied to is one follow-up, not a topic of one", () =>
            {
                var rows = TodoTopics.Group(new List<TodoItem>
                {
                    Item("Report", "e1", t0)
                }, null);

                var stats = TodoTopics.Summarize(rows);
                return rows.Count == 1 && !rows[0].IsTopic
                    && rows[0].Items.Count == 1
                    && stats.Items == 1 && stats.Topics == 0;
            });

            failed += Check("a title that cannot carry a topic always stands alone", () =>
            {
                var rows = TodoTopics.Group(new List<TodoItem>
                {
                    Item(string.Empty, "e1", t0),
                    Item(string.Empty, "e2", t0),
                    Item("Report", "e3", t0)
                }, null);

                var stats = TodoTopics.Summarize(rows);
                return rows.Count == 3 && stats.Items == 3 && stats.Topics == 0;
            });

            failed += Check("the order the user picked decides the lead and the order inside", () =>
            {
                var newest = Item("Report", "e1", t0);
                var older = Item("RE: Report", "e2", t0.AddDays(-1));
                var oldest = Item("RE: RE: Report", "e3", t0.AddDays(-3));

                var ordered = TodoOrdering.Sort(
                    new List<TodoItem> { older, oldest, newest }, TodoSortMode.NewestFirst, t0);
                var rows = TodoTopics.Group(ordered, null);

                return rows.Count == 1
                    && ReferenceEquals(rows[0].Lead, newest)
                    && ReferenceEquals(rows[0].Items[0], newest)
                    && ReferenceEquals(rows[0].Newest, newest)
                    && ReferenceEquals(rows[0].Oldest, oldest)
                    && rows[0].Items.Count == 3;
            });

            failed += Check("a finished follow-up still belongs to its topic, and is not counted open", () =>
            {
                var open = Item("Report", "e1", t0);
                var done = Item("RE: Report", "e2", t0.AddDays(-1));
                done.Status = TodoStatus.Done;

                var rows = TodoTopics.Group(new List<TodoItem> { open, done }, null);

                return rows.Count == 1 && rows[0].IsTopic
                    && rows[0].Items.Count == 2 && rows[0].OpenCount == 1;
            });

            failed += Check("the fold never loses or doubles a follow-up", () =>
            {
                var rows = TodoTopics.Group(new List<TodoItem>
                {
                    Item("Report", "e1", t0),
                    Item("RE: Report", "e2", t0),
                    Item("Invoice", "e3", t0),
                    Item("RE: Invoice", "e4", t0),
                    Item("Standalone", "e5", t0)
                }, null);

                var seen = 0;
                foreach (var row in rows)
                {
                    seen += row.Items.Count;
                }

                var stats = TodoTopics.Summarize(rows);
                return rows.Count == 3 && seen == 5 && stats.Items == 5 && stats.Topics == 2;
            });

            failed += Check("a list with nothing grouped reports no groupings", () =>
            {
                var rows = TodoTopics.Group(new List<TodoItem>
                {
                    Item("Report", "e1", t0),
                    Item("Invoice", "e2", t0),
                    Item("Standalone", "e3", t0)
                }, null);

                var stats = TodoTopics.Summarize(rows);
                return rows.Count == 3 && stats.Items == 3 && stats.Topics == 0 && stats.Done == 0;
            });

            failed += Check("the tally counts finished follow-ups once each", () =>
            {
                var a = Item("Report", "e1", t0);
                var b = Item("RE: Report", "e2", t0);
                var c = Item("Invoice", "e3", t0);
                a.Status = TodoStatus.Done;
                c.Status = TodoStatus.Done;

                var stats = TodoTopics.Summarize(TodoTopics.Group(
                    new List<TodoItem> { a, b, c }, null));

                return stats.Items == 3 && stats.Done == 2 && stats.Topics == 1;
            });

            failed += Check("an open topic is remembered as open, a closed one as closed", () =>
            {
                var items = new List<TodoItem>
                {
                    Item("Report", "e1", t0),
                    Item("RE: Report", "e2", t0.AddDays(-1))
                };

                var open = TodoTopics.Group(items, key => true);
                var closed = TodoTopics.Group(items, key => false);
                var alone = TodoTopics.Group(items, null);

                return open.Count == 1 && open[0].Expanded
                    && closed.Count == 1 && !closed[0].Expanded
                    && alone.Count == 1 && !alone[0].Expanded;
            });

            failed += Check("folding nothing is not a crash", () =>
            {
                var empty = TodoTopics.Group(new List<TodoItem>(), null);
                var stats = TodoTopics.Summarize(empty);
                var none = TodoTopics.Summarize(null);
                return empty.Count == 0 && stats.Items == 0 && stats.Topics == 0
                    && none.Items == 0 && none.Topics == 0;
            });

            return failed;
        }

        private static TodoItem Item(string title, string id, DateTime receivedUtc)
        {
            var item = TodoItem.Create(title, id, null, "sender@example.com", receivedUtc);
            item.CreatedOn = receivedUtc;
            item.UpdatedOn = receivedUtc;
            return item;
        }

        private static int Check(string name, Func<bool> body)
        {
            try
            {
                if (body())
                {
                    Console.WriteLine("PASS " + name);
                    return 0;
                }

                Console.WriteLine("FAIL " + name);
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL " + name + " :: " + ex.Message);
                return 1;
            }
        }
    }
}
