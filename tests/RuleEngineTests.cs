using System;
using System.Collections.Generic;
using System.Linq;
using OutlookAiHelper.Core.Classification;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Tests
{
    public static class RuleEngineTests
    {
        public static int Run()
        {
            var failed = 0;
            failed += Check("urgent keyword in subject", () =>
            {
                var engine = new RuleEngine(RuleOptions.CreateDefault());
                var result = engine.Classify(Mail("1", "今天請盡快回覆", false), null);
                return result.Quadrant == Quadrant.Q2UrgentNotImportant
                    || result.Quadrant == Quadrant.Q1UrgentImportant;
            });

            failed += Check("important keyword in subject", () =>
            {
                var engine = new RuleEngine(RuleOptions.CreateDefault());
                var result = engine.Classify(Mail("2", "專案報價確認", true), null);
                return result.Quadrant == Quadrant.Q3ImportantNotUrgent
                    || result.Quadrant == Quadrant.Q1UrgentImportant;
            });

            failed += Check("urgent + important => Q1", () =>
            {
                var engine = new RuleEngine(RuleOptions.CreateDefault());
                var result = engine.Classify(Mail("3", "緊急專案付款", true), null);
                return result.Quadrant == Quadrant.Q1UrgentImportant;
            });

            failed += Check("no signal => Q4", () =>
            {
                var engine = new RuleEngine(RuleOptions.CreateDefault());
                var result = engine.Classify(Mail("4", "newsletter", true), null);
                return result.Quadrant == Quadrant.Q4Neither && result.Reasons.Count > 0;
            });

            failed += Check("override wins", () =>
            {
                var engine = new RuleEngine(RuleOptions.CreateDefault());
                var map = new Dictionary<string, OverrideEntry>
                {
                    { "5", OverrideEntry.Manual("5", Quadrant.Q4Neither) }
                };
                var result = engine.Classify(Mail("5", "緊急專案", true), map);
                return result.FromOverride && result.Quadrant == Quadrant.Q4Neither;
            });

            failed += Check("vip adds importance", () =>
            {
                var options = RuleOptions.CreateDefault();
                options.VipAddresses.Add("boss@contoso.com");
                var engine = new RuleEngine(options);
                var mail = Mail("6", "hello", true);
                mail.FromAddress = "boss@contoso.com";
                var result = engine.Classify(mail, null);
                return result.ImportanceScore >= 3 && result.Quadrant == Quadrant.Q3ImportantNotUrgent;
            });

            failed += Check("reasons explain keyword hit", () =>
            {
                var engine = new RuleEngine(RuleOptions.CreateDefault());
                var result = engine.Classify(Mail("7", "ASAP contract", true), null);
                return result.Reasons.Any(r => r.Code == "keyword.urgent" && r.Detail == "ASAP")
                    && result.Reasons.Any(r => r.Code == "keyword.important");
            });

            failed += Check("null mail is safe", () =>
            {
                var engine = new RuleEngine(RuleOptions.CreateDefault());
                var result = engine.Classify(null, null);
                return result.Quadrant == Quadrant.Q4Neither;
            });

            return failed;
        }

        private static MailSummary Mail(string id, string subject, bool read)
        {
            return new MailSummary
            {
                EntryId = id,
                Subject = subject,
                FromName = "Test",
                FromAddress = "test@contoso.com",
                ReceivedOn = DateTime.Now,
                IsRead = read,
                HasFlag = false,
                IsMeetingRequest = false,
                Importance = 1,
                FolderPath = "Inbox"
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
