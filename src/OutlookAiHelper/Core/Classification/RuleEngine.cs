using System.Collections.Generic;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Core.Classification
{
    public sealed class RuleEngine : IClassifier
    {
        private readonly RuleOptions _options;

        public RuleEngine(RuleOptions options)
        {
            _options = options ?? RuleOptions.CreateDefault();
        }

        public ClassificationResult Classify(MailSummary mail, IDictionary<string, OverrideEntry> overrides)
        {
            if (mail == null)
            {
                return new ClassificationResult(
                    string.Empty,
                    Quadrant.Q4Neither,
                    0,
                    0,
                    new List<ScoreReason>(),
                    false);
            }

            OverrideEntry manual;
            if (overrides != null && !string.IsNullOrEmpty(mail.EntryId) && overrides.TryGetValue(mail.EntryId, out manual) && manual != null)
            {
                var overrideReasons = new List<ScoreReason>
                {
                    new ScoreReason("override.manual", "reason.override.manual", 0, "Manual")
                };
                return new ClassificationResult(mail.EntryId, manual.Quadrant, 0, 0, overrideReasons, true);
            }

            var reasons = new List<ScoreReason>();
            var urgency = 0;
            var importance = 0;
            var subject = mail.Subject ?? string.Empty;

            string hit;
            if (RuleOptions.ContainsKeyword(subject, _options.UrgentKeywords, out hit))
            {
                urgency = Cap(urgency + 2, _options.MaxUrgencyScore);
                reasons.Add(new ScoreReason("keyword.urgent", "reason.keyword.urgent", 2, hit));
            }

            if (RuleOptions.ContainsKeyword(subject, _options.ImportantKeywords, out hit))
            {
                importance = Cap(importance + 2, _options.MaxImportanceScore);
                reasons.Add(new ScoreReason("keyword.important", "reason.keyword.important", 2, hit));
            }

            if (mail.IsMeetingRequest)
            {
                urgency = Cap(urgency + 2, _options.MaxUrgencyScore);
                reasons.Add(new ScoreReason("signal.meeting", "reason.signal.meeting", 2, "MeetingRequest"));
            }

            if (!mail.IsRead)
            {
                urgency = Cap(urgency + 1, _options.MaxUrgencyScore);
                reasons.Add(new ScoreReason("signal.unread", "reason.signal.unread", 1, "Unread"));
            }

            if (mail.HasFlag)
            {
                importance = Cap(importance + 1, _options.MaxImportanceScore);
                reasons.Add(new ScoreReason("signal.flag", "reason.signal.flag", 1, "Flag"));
            }

            if (mail.Importance >= 2)
            {
                importance = Cap(importance + 2, _options.MaxImportanceScore);
                reasons.Add(new ScoreReason("signal.importance", "reason.signal.importance", 2, "High"));
            }

            if (RuleOptions.IsVip(mail.FromAddress, _options.VipAddresses))
            {
                importance = Cap(importance + 3, _options.MaxImportanceScore);
                reasons.Add(new ScoreReason("signal.vip", "reason.signal.vip", 3, mail.FromAddress));
            }

            if (reasons.Count == 0)
            {
                reasons.Add(new ScoreReason("signal.default", "reason.signal.default", 0, "NoSignal"));
            }

            var isUrgent = urgency >= _options.UrgentThreshold;
            var isImportant = importance >= _options.ImportantThreshold;
            var quadrant = ResolveQuadrant(isUrgent, isImportant);
            return new ClassificationResult(mail.EntryId, quadrant, urgency, importance, reasons, false);
        }

        public static Quadrant ResolveQuadrant(bool urgent, bool important)
        {
            if (urgent && important)
            {
                return Quadrant.Q1UrgentImportant;
            }

            if (urgent)
            {
                return Quadrant.Q2UrgentNotImportant;
            }

            if (important)
            {
                return Quadrant.Q3ImportantNotUrgent;
            }

            return Quadrant.Q4Neither;
        }

        private static int Cap(int value, int max)
        {
            return value > max ? max : value;
        }
    }
}
