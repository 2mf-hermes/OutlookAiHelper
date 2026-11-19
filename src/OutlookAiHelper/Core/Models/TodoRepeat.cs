using System;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// How a follow-up comes back after it is ticked off. Stored as a name, not a
    /// number, so an older build reading a newer file still understands "Weekly" and
    /// an unknown value degrades to <see cref="None"/> instead of shifting every item.
    /// </summary>
    public enum TodoRepeat
    {
        None = 0,
        Daily = 1,
        Weekly = 2,
        Monthly = 3,
        Yearly = 4
    }

    /// <summary>
    /// The date arithmetic behind repeating follow-ups. Kept here - not in the UI - so
    /// the tests can pin the month-end rule without a window: 1 月 31 日 repeats on
    /// 2 月 28/29 日, then returns to 3 月 31 日 because the item remembers the day it
    /// was anchored on (<see cref="TodoItem.RepeatAnchorDay"/>), which is what the
    /// mainstream task apps do.
    /// </summary>
    public static class TodoRecurrence
    {
        public static bool IsRepeating(TodoRepeat repeat)
        {
            return repeat != TodoRepeat.None;
        }

        /// <summary>
        /// The next occurrence of <paramref name="fromUtc"/>. <paramref name="anchorDay"/>
        /// is the day of the month the series was created on (0 means "use fromUtc.Day").
        /// The time of day is preserved; the kind (UTC) is preserved.
        /// </summary>
        public static DateTime Next(DateTime fromUtc, TodoRepeat repeat, int anchorDay = 0)
        {
            switch (repeat)
            {
                case TodoRepeat.Daily:
                    return fromUtc.AddDays(1);
                case TodoRepeat.Weekly:
                    return fromUtc.AddDays(7);
                case TodoRepeat.Monthly:
                    return AddMonthsOnDay(fromUtc, 1, anchorDay);
                case TodoRepeat.Yearly:
                    return AddMonthsOnDay(fromUtc, 12, anchorDay);
                default:
                    return fromUtc;
            }
        }

        /// <summary>Stable key for the settings-free JSON record.</summary>
        public static string Key(TodoRepeat repeat)
        {
            switch (repeat)
            {
                case TodoRepeat.Daily:
                    return "Daily";
                case TodoRepeat.Weekly:
                    return "Weekly";
                case TodoRepeat.Monthly:
                    return "Monthly";
                case TodoRepeat.Yearly:
                    return "Yearly";
                default:
                    return "None";
            }
        }

        public static TodoRepeat Parse(string value)
        {
            if (string.Equals(value, "Daily", StringComparison.OrdinalIgnoreCase))
            {
                return TodoRepeat.Daily;
            }

            if (string.Equals(value, "Weekly", StringComparison.OrdinalIgnoreCase))
            {
                return TodoRepeat.Weekly;
            }

            if (string.Equals(value, "Monthly", StringComparison.OrdinalIgnoreCase))
            {
                return TodoRepeat.Monthly;
            }

            if (string.Equals(value, "Yearly", StringComparison.OrdinalIgnoreCase))
            {
                return TodoRepeat.Yearly;
            }

            return TodoRepeat.None;
        }

        private static DateTime AddMonthsOnDay(DateTime from, int months, int anchorDay)
        {
            var firstOfMonth = new DateTime(from.Year, from.Month, 1, from.Hour, from.Minute, from.Second, from.Kind);
            var target = firstOfMonth.AddMonths(months);
            var day = anchorDay > 0 ? anchorDay : from.Day;
            var lastDay = DateTime.DaysInMonth(target.Year, target.Month);
            if (day > lastDay)
            {
                day = lastDay;
            }

            return new DateTime(target.Year, target.Month, day, target.Hour, target.Minute, target.Second, target.Kind);
        }
    }
}
