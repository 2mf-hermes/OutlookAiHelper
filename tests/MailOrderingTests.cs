using System;
using System.Collections.Generic;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Classification;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// Covers the one mail order shared by the mail list and the mails listed beside a
    /// follow-up: which mode a stored name reads as, and what each mode does to a list. The
    /// rows themselves are drawn by the window, so what is asserted here is the order, not the
    /// pixels.
    /// </summary>
    public static class MailOrderingTests
    {
        public static int Run()
        {
            var failed = 0;

            failed += Check("the default mode is newest first and it is the first choice", () =>
            {
                return MailOrdering.DefaultMode == MailSortMode.NewestFirst
                    && MailOrdering.Modes.Length > 0
                    && MailOrdering.Modes[0] == MailSortMode.NewestFirst;
            });

            failed += Check("every mode has a label and the label keys are distinct", () =>
            {
                var seen = new List<string>();
                foreach (var mode in MailOrdering.Modes)
                {
                    var key = MailOrdering.LabelKey(mode);
                    if (string.IsNullOrEmpty(key) || seen.Contains(key))
                    {
                        return false;
                    }

                    seen.Add(key);
                }

                return seen.Count == MailOrdering.Modes.Length;
            });

            failed += Check("a mode survives being written down and read back", () =>
            {
                foreach (var mode in MailOrdering.Modes)
                {
                    if (MailOrdering.ParseMode(mode.ToString()) != mode)
                    {
                        return false;
                    }
                }

                return true;
            });

            failed += Check("an unknown or missing stored name reads as the default", () =>
            {
                return MailOrdering.ParseMode(null) == MailOrdering.DefaultMode
                    && MailOrdering.ParseMode(string.Empty) == MailOrdering.DefaultMode
                    && MailOrdering.ParseMode("whatever") == MailOrdering.DefaultMode;
            });

            failed += Check("newest first orders a list by time, latest mail first", () =>
            {
                var ordered = MailOrdering.SortMails(
                    new[]
                    {
                        Mail("e1", "Q3 預算", "Amy", "amy@contoso.com", 0),
                        Mail("e2", "Q3 預算", "Amy", "amy@contoso.com", 30),
                        Mail("e3", "Q3 預算", "Amy", "amy@contoso.com", 10)
                    },
                    MailSortMode.NewestFirst);

                return ordered[0].Mail.EntryId == "e2"
                    && ordered[1].Mail.EntryId == "e3"
                    && ordered[2].Mail.EntryId == "e1";
            });

            failed += Check("oldest first is the same list the other way up", () =>
            {
                var ordered = MailOrdering.SortMails(
                    new[]
                    {
                        Mail("e1", "Q3 預算", "Amy", "amy@contoso.com", 0),
                        Mail("e2", "Q3 預算", "Amy", "amy@contoso.com", 30),
                        Mail("e3", "Q3 預算", "Amy", "amy@contoso.com", 10)
                    },
                    MailSortMode.OldestFirst);

                return ordered[0].Mail.EntryId == "e1"
                    && ordered[1].Mail.EntryId == "e3"
                    && ordered[2].Mail.EntryId == "e2";
            });

            failed += Check("a mail with no summary behind it is dropped, not ordered", () =>
            {
                var ordered = MailOrdering.SortMails(
                    new[]
                    {
                        Mail("e1", "Q3 預算", "Amy", "amy@contoso.com", 0),
                        null,
                        new ScanResultItem { Mail = null },
                        Mail("e2", "Q3 預算", "Amy", "amy@contoso.com", 5)
                    },
                    MailSortMode.NewestFirst);

                return ordered.Count == 2 && ordered[0].Mail.EntryId == "e2";
            });

            failed += Check("the sender mode groups one sender together, newest first inside", () =>
            {
                // One subject throughout, so what the groups are made of is only the sender and
                // the expected order inside a group can only come from the times.
                var ordered = MailOrdering.SortMails(
                    new[]
                    {
                        Mail("e1", "Q3 預算", "Ben", "ben@contoso.com", 10),
                        Mail("e2", "Q3 預算", "Amy", "amy@contoso.com", 0),
                        Mail("e3", "Q3 預算", "Amy", "amy@contoso.com", 20),
                        Mail("e4", "Q3 預算", "Ben", "ben@contoso.com", 30)
                    },
                    MailSortMode.Sender);

                return ordered[0].Mail.EntryId == "e3"
                    && ordered[1].Mail.EntryId == "e2"
                    && ordered[2].Mail.EntryId == "e4"
                    && ordered[3].Mail.EntryId == "e1";
            });

            failed += Check("a topic leads with its newest mail when the list is newest first", () =>
            {
                // e1 is the scan's highest-scoring mail and the oldest of the topic; under
                // newest first the row has to present e3 and the row itself has to sit at e3's
                // time. This is the reported defect, pinned: the row used to show e1.
                var rows = MailOrdering.BuildRows(
                    new[]
                    {
                        Mail("e1", "Q3 預算", "Amy", "amy@contoso.com", 0, 90),
                        Mail("e2", "Re: Q3 預算", "Amy", "amy@contoso.com", 10, 10),
                        Mail("e3", "Re: Re: Q3 預算", "Amy", "amy@contoso.com", 60, 20),
                        Mail("e4", "旅遊補助", "Ben", "ben@contoso.com", 30, 50)
                    },
                    MailSortMode.NewestFirst,
                    null,
                    null);

                return rows.Count == 2
                    && rows[0].Head.Mail.EntryId == "e3"
                    && rows[0].Newest.Mail.EntryId == "e3"
                    && rows[0].Mails[0].Mail.EntryId == "e3"
                    && rows[1].Head.Mail.EntryId == "e4";
            });

            failed += Check("inside a topic no mail is listed above a newer one, newest first", () =>
            {
                var rows = MailOrdering.BuildRows(
                    new[]
                    {
                        Mail("e1", "Q3 預算", "Amy", "amy@contoso.com", 0),
                        Mail("e2", "Re: Q3 預算", "Amy", "amy@contoso.com", 30),
                        Mail("e3", "Re: Re: Q3 預算", "Amy", "amy@contoso.com", 10)
                    },
                    MailSortMode.NewestFirst,
                    null,
                    null);

                return rows.Count == 1
                    && rows[0].Mails[0].Mail.EntryId == "e2"
                    && rows[0].Mails[1].Mail.EntryId == "e3"
                    && rows[0].Mails[2].Mail.EntryId == "e1";
            });

            failed += Check("a mode that is not about time still decides which mail a topic presents", () =>
            {
                // Quadrant ranks the older Q1 mail above the newer Q3 one, so the row must
                // present e1 there, while newest first must present e2. Fixing one mode must
                // not decide the other.
                var quiet = new ScanResultItem
                {
                    Mail = Summary("Q3 預算", "Amy", "amy@contoso.com", "e1", 0),
                    Classification = new ClassificationResult(
                        "e1", Quadrant.Q1UrgentImportant, 90, 90, new List<ScoreReason>(), false)
                };
                var later = new ScanResultItem
                {
                    Mail = Summary("Re: Q3 預算", "Amy", "amy@contoso.com", "e2", 60),
                    Classification = new ClassificationResult(
                        "e2", Quadrant.Q3ImportantNotUrgent, 20, 80, new List<ScoreReason>(), false)
                };

                var byQuadrant = MailOrdering.BuildRows(
                    new[] { quiet, later }, MailSortMode.Quadrant, null, null);
                var byNewest = MailOrdering.BuildRows(
                    new[] { quiet, later }, MailSortMode.NewestFirst, null, null);

                return byQuadrant.Count == 1
                    && byQuadrant[0].Head.Mail.EntryId == "e1"
                    && byNewest.Count == 1
                    && byNewest[0].Head.Mail.EntryId == "e2";
            });

            failed += Check("the same mails ordered by quadrant put the urgent ones first", () =>
            {
                var q1 = new ScanResultItem
                {
                    Mail = Summary("主題甲", "Amy", "amy@contoso.com", "e1", 0),
                    Classification = new ClassificationResult(
                        "e1", Quadrant.Q1UrgentImportant, 90, 90, new List<ScoreReason>(), false)
                };
                var q3 = new ScanResultItem
                {
                    Mail = Summary("主題乙", "Amy", "amy@contoso.com", "e2", 10),
                    Classification = new ClassificationResult(
                        "e2", Quadrant.Q3ImportantNotUrgent, 30, 80, new List<ScoreReason>(), false)
                };
                var q4 = new ScanResultItem
                {
                    Mail = Summary("主題丙", "Amy", "amy@contoso.com", "e3", 20),
                    Classification = new ClassificationResult(
                        "e3", Quadrant.Q4Neither, 10, 10, new List<ScoreReason>(), false)
                };

                var ordered = MailOrdering.SortMails(
                    new[] { q4, q3, q1 },
                    MailSortMode.Quadrant);

                return ordered[0].Mail.EntryId == "e1"
                    && ordered[1].Mail.EntryId == "e2"
                    && ordered[2].Mail.EntryId == "e3";
            });

            return failed;
        }

        private static ScanResultItem Mail(
            string entryId,
            string subject,
            string from,
            string address,
            int minutesLater,
            int urgency = 60)
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
