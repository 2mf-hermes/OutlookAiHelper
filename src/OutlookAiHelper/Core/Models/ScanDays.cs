using System;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// The bounds and parsing rule for the scan range. The toolbar's day picker offers
    /// four presets and a custom entry; both the custom dialog and the settings page
    /// go through <see cref="TryParse"/>, so "0", "-3", "abc" and "9999" cannot reach
    /// the reader as a scan window.
    /// </summary>
    public static class ScanDays
    {
        public const int Min = 1;
        public const int Max = 365;
        public const int Default = 30;

        /// <summary>The presets the toolbar shows, in order.</summary>
        public static readonly int[] Presets = { 7, 14, 30, 60, 90 };

        public static bool TryParse(string text, out int days)
        {
            days = Default;
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            int parsed;
            if (!int.TryParse(text.Trim(), out parsed))
            {
                return false;
            }

            if (parsed < Min || parsed > Max)
            {
                return false;
            }

            days = parsed;
            return true;
        }

        public static int Clamp(int days)
        {
            if (days < Min)
            {
                return Min;
            }

            if (days > Max)
            {
                return Max;
            }

            return days;
        }
    }
}
