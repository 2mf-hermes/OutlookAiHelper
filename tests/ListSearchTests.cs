using System;
using System.Collections.Generic;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// Covers the search both lists share. What is asserted here is which rows survive a query,
    /// not what the window draws: the fields each page feeds in are its own business, the rule for
    /// answering a query is not. A query that asks for nothing has to match everything, because a
    /// search field left empty is the list unfiltered, not the list emptied.
    /// </summary>
    public static class ListSearchTests
    {
        public static int Run()
        {
            var failed = 0;

            failed += Check("an empty query asks for nothing and matches every row", () =>
            {
                return ListSearch.Matches(null, "anything")
                    && ListSearch.Matches(string.Empty, "anything")
                    && ListSearch.Matches(null, (string[])null)
                    && !ListSearch.IsActive(null)
                    && !ListSearch.IsActive(string.Empty);
            });

            failed += Check("a query of spaces alone asks for nothing", () =>
            {
                return ListSearch.Tokens("   ").Length == 0
                    && ListSearch.Tokens("\t \u3000").Length == 0
                    && ListSearch.Matches("   ", "anything");
            });

            failed += Check("every word of a query is lower-cased and whitespace is what splits it", () =>
            {
                var words = ListSearch.Tokens("  Wang\tINVOICE \u3000friday ");
                return words.Length == 3
                    && words[0] == "wang"
                    && words[1] == "invoice"
                    && words[2] == "friday";
            });

            failed += Check("a word finds a mail by its subject", () =>
            {
                return ListSearch.Matches("invoice", Mail().Subject, Mail().Preview, Mail().FromName, Mail().FromAddress, Mail().ToLine, Mail().CcLine);
            });

            failed += Check("a mail is found by the sender's name and by the sender's address", () =>
            {
                var mail = Mail();
                return ListSearch.Matches("wang", mail.FromName, mail.FromAddress)
                    && ListSearch.Matches("example.com", mail.FromName, mail.FromAddress);
            });

            failed += Check("a mail is found by a recipient, in either the To or the Cc line", () =>
            {
                var mail = Mail();
                return ListSearch.Matches("finance", mail.ToLine, mail.CcLine)
                    && ListSearch.Matches("auditor", mail.ToLine, mail.CcLine)
                    && !ListSearch.Matches("auditor", mail.ToLine);
            });

            failed += Check("a mail is found by a word from the opening of its text", () =>
            {
                return ListSearch.Matches("renewal", Mail().Preview);
            });

            failed += Check("every word has to be found, in any order and in any fields", () =>
            {
                var mail = Mail();
                var fields = new[] { mail.Subject, mail.Preview, mail.FromName, mail.FromAddress, mail.ToLine, mail.CcLine };
                return ListSearch.Matches("invoice wang", fields)
                    && ListSearch.Matches("wang invoice", fields)
                    && ListSearch.Matches("wang finance", fields)
                    && !ListSearch.Matches("invoice missing", fields);
            });

            failed += Check("a word that is nowhere in the row finds nothing", () =>
            {
                var mail = Mail();
                return !ListSearch.Matches("unheard-of-word", mail.Subject, mail.Preview, mail.FromName, mail.FromAddress, mail.ToLine, mail.CcLine);
            });

            failed += Check("case does not decide a match, and a half-typed word already narrows", () =>
            {
                return ListSearch.Matches("INVOICE", "Quarterly invoice")
                    && ListSearch.Matches("invo", "Quarterly invoice")
                    && ListSearch.Matches("quarterly", "QUARTERLY INVOICE");
            });

            failed += Check("a null field is skipped rather than swallowing the query", () =>
            {
                var mail = Mail();
                return ListSearch.Matches("invoice", null, mail.Subject, null)
                    && !ListSearch.Matches("auditor", null, mail.Subject, null);
            });

            return failed;
        }

        /// <summary>
        /// One mail with something searchable in every field the mail list offers, so a test that
        /// passes proves the field it names is really the one that answered.
        /// </summary>
        private static MailSummary Mail()
        {
            return new MailSummary
            {
                EntryId = "e1",
                Subject = "Quarterly invoice for the Kaohsiung site",
                FromName = "Wang Mei-ling",
                FromAddress = "mei-ling.wang@example.com",
                ToLine = "Finance Team <finance@example.com>",
                CcLine = "External Auditor <auditor@example.org>",
                Preview = "Please confirm the renewal figures before Friday.",
                ReceivedOn = new DateTime(2026, 9, 20, 9, 0, 0, DateTimeKind.Utc)
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
