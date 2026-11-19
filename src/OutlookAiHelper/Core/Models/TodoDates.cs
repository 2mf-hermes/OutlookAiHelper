using System;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// The one place that decides what a picked calendar day means. A deadline is the
    /// end of that local day, not midnight at its start, so a follow-up due "today" is
    /// not already overdue the moment it is created. The result is a UTC instant, the
    /// same currency the store and the overdue test use.
    /// </summary>
    public static class TodoDates
    {
        /// <summary>
        /// The last tick of <paramref name="day"/> in local time, as UTC. Only the date
        /// part of the argument is used; the time of day the caller passed is ignored.
        /// </summary>
        public static DateTime EndOfLocalDay(DateTime day)
        {
            var start = day.Date;
            var end = start.AddDays(1).AddTicks(-1);
            return DateTime.SpecifyKind(end, DateTimeKind.Local).ToUniversalTime();
        }

        /// <summary>The local calendar day a stored instant falls on.</summary>
        public static DateTime LocalDayOf(DateTime utc)
        {
            var local = utc.Kind == DateTimeKind.Utc ? utc.ToLocalTime() : utc;
            return local.Date;
        }

        /// <summary>
        /// A picked day as a stored value, or null when the picker was cleared. Kept
        /// here so the composer and any future surface agree on the round trip.
        /// </summary>
        public static DateTime? DueFromDay(DateTime? day)
        {
            return day.HasValue ? (DateTime?)EndOfLocalDay(day.Value) : null;
        }

        /// <summary>The local day of a stored due instant, for a date picker to show.</summary>
        public static DateTime? DayFromDue(DateTime? utc)
        {
            return utc.HasValue ? (DateTime?)LocalDayOf(utc.Value) : null;
        }

        /// <summary>
        /// Normalises a time the mail reader handed back. Outlook reports local time with
        /// an unspecified Kind, a cloud reader may already say Utc, and a missing value
        /// comes back as <c>default</c>; all three end up as the same UTC currency the
        /// store and the overdue test use. Null when there is no time at all.
        /// </summary>
        public static DateTime? ToUtc(DateTime value)
        {
            if (value == default(DateTime))
            {
                return null;
            }

            return value.Kind == DateTimeKind.Utc ? value : value.ToUniversalTime();
        }
    }
}
