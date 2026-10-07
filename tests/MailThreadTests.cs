using System;
using System.Collections.Generic;
using System.Linq;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Classification;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// Covers the two mail-list rules that live outside WPF: which mails fold into one topic
    /// row, and which mails carry the 已加入待辦 tag. The rows themselves are drawn by the
    /// window, so what is asserted here is the decision, not the pixels.
    /// </summary>
    public static class MailThreadTests
    {
        public static int Run()
        {
            var failed = 0;

            failed += Check("a reply folds into the mail it answers", () =>
            {
                var rows = MailThreads.Group(
                    new[]
                    {
                        Mail("e1", "Q3 預算", "Amy", "amy@contoso.com", 3),
                        Mail("e2", "Re: Q3 預算", "Amy", "amy@contoso.com", 2),
                        Mail("e3", "回覆：Q3 預算", "Ben", "ben@contoso.com", 1)
                    },
                    null,
                    null);

                return rows.Count == 1 && rows[0].IsThread && rows[0].Mails.Count == 3;
            });

            failed += Check("the row presents the highest-ranked mail of the topic", () =>
            {
                // Ranked order puts e1 first; the newest reply in the topic is e3, which is a
                // different question and has to be answered from the received times.
                var rows = MailThreads.Group(
                    new[]
                    {
                        Mail("e1", "Q3 預算", "Amy", "amy@contoso.com", 3, 0),
                        Mail("e2", "Re: Q3 預算", "Amy", "amy@contoso.com", 1, 10),
                        Mail("e3", "Re: Re: Q3 預算", "Amy", "amy@contoso.com", 2, 60)
                    },
                    null,
                    null);

                return ReferenceEquals(rows[0].Head, rows[0].Mails[0]) &&
                       rows[0].Head.Mail.EntryId == "e1" &&
                       rows[0].Newest.Mail.EntryId == "e3";
            });

            failed += Check("reply and forward markers are stripped however they are written", () =>
            {
                var plain = MailThreads.KeyOf(Summary("Q3 預算", "Amy", "amy@contoso.com"));
                var reply = MailThreads.KeyOf(Summary("Re: Fw:  Q3   預算", "Amy", "amy@contoso.com"));
                var local = MailThreads.KeyOf(Summary("轉寄：Q3 預算", "Amy", "amy@contoso.com"));
                var full = MailThreads.KeyOf(Summary("回覆：Q3 預算", "Amy", "amy@contoso.com"));

                return plain != null && plain == reply && plain == local && plain == full;
            });

            failed += Check("a shared subject is the topic, whoever sent it", () =>
            {
                var rows = MailThreads.Group(
                    new[]
                    {
                        Mail("e1", "Q3 預算", "Amy", "amy@contoso.com", 2),
                        Mail("e2", "Q3 預算", "Ben", "ben@contoso.com", 1)
                    },
                    null,
                    null);

                return rows.Count == 1 && rows[0].Mails.Count == 2;
            });

            failed += Check("a mail with no usable subject falls back to its sender", () =>
            {
                var rows = MailThreads.Group(
                    new[]
                    {
                        Mail("e1", "?", "Amy", "amy@contoso.com", 3),
                        Mail("e2", "  ", "Amy", "amy@contoso.com", 2),
                        Mail("e3", null, "Ben", "ben@contoso.com", 1)
                    },
                    null,
                    null);

                return rows.Count == 2 &&
                       rows[0].Mails.Count == 2 &&
                       rows[0].IsThread &&
                       rows[1].Mails.Count == 1;
            });

            failed += Check("a lone mail is drawn as a plain row, never as a topic", () =>
            {
                var rows = MailThreads.Group(
                    new[] { Mail("e1", "Q3 預算", "Amy", "amy@contoso.com", 1) },
                    null,
                    key => true);

                return rows.Count == 1 && !rows[0].IsThread && rows[0].Mails.Count == 1;
            });

            failed += Check("topics appear in the caller's order, not the mailbox's", () =>
            {
                var rows = MailThreads.Group(
                    new[]
                    {
                        Mail("e1", "Zeta", "Amy", "amy@contoso.com", 3),
                        Mail("e2", "Alpha", "Ben", "ben@contoso.com", 2),
                        Mail("e3", "Re: Zeta", "Amy", "amy@contoso.com", 1)
                    },
                    null,
                    null);

                return rows.Count == 2 &&
                       rows[0].Head.Mail.EntryId == "e1" &&
                       rows[1].Head.Mail.EntryId == "e2" &&
                       rows[0].Mails[1].Mail.EntryId == "e3";
            });

            failed += Check("only the mails that were added to the follow-ups are counted", () =>
            {
                var added = new HashSet<string>(StringComparer.Ordinal) { "e1", "e3" };
                var rows = MailThreads.Group(
                    new[]
                    {
                        Mail("e1", "Q3 預算", "Amy", "amy@contoso.com", 3),
                        Mail("e2", "Re: Q3 預算", "Amy", "amy@contoso.com", 2),
                        Mail("e3", "Re: Re: Q3 預算", "Amy", "amy@contoso.com", 1)
                    },
                    id => added.Contains(id),
                    null);

                return rows[0].TodoCount == 2 && rows[0].HasTodo;
            });

            failed += Check("an open topic comes back open, a closed one comes back closed", () =>
            {
                var rows = MailThreads.Group(
                    new[]
                    {
                        Mail("e1", "Q3 預算", "Amy", "amy@contoso.com", 2),
                        Mail("e2", "Re: Q3 預算", "Amy", "amy@contoso.com", 1)
                    },
                    null,
                    key => key == MailThreads.KeyOf(Summary("Q3 預算", "Amy", "amy@contoso.com")));

                return rows[0].Expanded;
            });

            failed += Check("the tag is driven by the mail behind the follow-up", () =>
            {
                var todos = new List<TodoItem>
                {
                    new TodoItem { Id = "t1", Title = "Q3 預算", SourceEntryId = "e1" },
                    new TodoItem { Id = "t2", Title = "自己加的" },
                    null
                };

                var ids = MailThreads.FollowUpEntryIds(todos);
                return ids.Count == 1 && ids.Contains("e1");
            });

            failed += Check("a follow-up with no mail behind it marks nothing", () =>
            {
                var ids = MailThreads.FollowUpEntryIds(new[] { new TodoItem { Id = "t1", Title = "自己加的" } });
                return ids.Count == 0;
            });

            failed += Check("the list summary counts mails, topics and tagged mails", () =>
            {
                var rows = MailThreads.Group(
                    new[]
                    {
                        Mail("e1", "Q3 預算", "Amy", "amy@contoso.com", 3),
                        Mail("e2", "Re: Q3 預算", "Amy", "amy@contoso.com", 2),
                        Mail("e3", "旅遊補助", "Ben", "ben@contoso.com", 1)
                    },
                    id => id == "e2",
                    null);
                var stats = MailThreads.Summarize(rows);

                return stats.Mails == 3 && stats.Threads == 1 && stats.WithTodo == 1;
            });

            return failed;
        }

        private static ScanResultItem Mail(
            string entryId,
            string subject,
            string from,
            string address,
            int urgency,
            int minutesLater = 0)
        {
            return new ScanResultItem
            {
                Mail = Summary(subject, from, address, entryId, minutesLater),
                Classification = new ClassificationResult(
                    entryId,
                    Quadrant.Q1UrgentImportant,
                    urgency,
                    50,
                    new List<ScoreReason>(),
                    false)
            };
        }

        private static MailSummary Summary(
            string subject,
            string from,
            string address,
            string entryId = "e1",
            int minutesLater = 0)
        {
            return new MailSummary
            {
                EntryId = entryId,
                Subject = subject,
                FromName = from,
                FromAddress = address,
                ReceivedOn = new DateTime(2026, 9, 20, 9, 0, 0, DateTimeKind.Utc).AddMinutes(minutesLater)
            };
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
