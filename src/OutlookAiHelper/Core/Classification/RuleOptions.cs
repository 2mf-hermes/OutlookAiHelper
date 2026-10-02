using System;
using System.Collections.Generic;

namespace OutlookAiHelper.Core.Classification
{
    public sealed class RuleOptions
    {
        public RuleOptions()
        {
            UrgentKeywords = DefaultUrgentKeywords();
            ImportantKeywords = DefaultImportantKeywords();
            VipAddresses = new List<string>();
            UrgentThreshold = 2;
            ImportantThreshold = 2;
            MaxUrgencyScore = 6;
            MaxImportanceScore = 6;
        }

        public IList<string> UrgentKeywords { get; set; }
        public IList<string> ImportantKeywords { get; set; }
        public IList<string> VipAddresses { get; set; }
        public int UrgentThreshold { get; set; }
        public int ImportantThreshold { get; set; }
        public int MaxUrgencyScore { get; set; }
        public int MaxImportanceScore { get; set; }

        public static RuleOptions CreateDefault()
        {
            return new RuleOptions();
        }

        public static List<string> DefaultUrgentKeywords()
        {
            return new List<string>
            {
                "緊急", "急件", "盡快", "立刻", "馬上", "今天", "今日",
                "ASAP", "urgent", "immediate", "deadline", "到期", "截止"
            };
        }

        public static List<string> DefaultImportantKeywords()
        {
            return new List<string>
            {
                "專案", "project", "合約", "契約", "報價", "付款", "發票",
                "客訴", "投訴", "簽核", "核准", "會議", "invoice", "contract",
                "quote", "approval", "complaint"
            };
        }

        public static bool ContainsKeyword(string text, IEnumerable<string> keywords, out string hit)
        {
            hit = null;
            if (string.IsNullOrEmpty(text) || keywords == null)
            {
                return false;
            }

            foreach (var keyword in keywords)
            {
                if (string.IsNullOrEmpty(keyword))
                {
                    continue;
                }

                if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    hit = keyword;
                    return true;
                }
            }

            return false;
        }

        public static bool IsVip(string fromAddress, IEnumerable<string> vipAddresses)
        {
            if (string.IsNullOrEmpty(fromAddress) || vipAddresses == null)
            {
                return false;
            }

            foreach (var vip in vipAddresses)
            {
                if (string.IsNullOrEmpty(vip))
                {
                    continue;
                }

                if (string.Equals(fromAddress.Trim(), vip.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
