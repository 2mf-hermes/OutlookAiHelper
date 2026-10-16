using System;
using System.Collections.Generic;
using System.Linq;
using OutlookAiHelper.Core.Classification;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// How the follow-up list is ordered. "Manual" is the order the store returned, which is
    /// the arrival order the page has always shown; NewestFirst says so out loud and is what
    /// the page opens with.
    /// </summary>
    public enum TodoSortMode
    {
        Manual,
        LongestWaiting,
        Quadrant,
        RecentlyUpdated,

        /// <summary>Soonest due first; items without a due date come last.</summary>
        DueSoonest,

        /// <summary>
        /// Newest arrival first. It is the order the page opens with, so it is also what an
        /// unrecognised stored preference falls back to.
        /// </summary>
        NewestFirst
    }

    /// <summary>
    /// Pure ordering rules for the follow-up list, kept out of the window so the
    /// "which one has been waiting longest" decision is testable without Outlook or UI.
    /// </summary>
    public static class TodoOrdering
    {
        /// <summary>Fallback rank for an item whose quadrant hint we cannot resolve.</summary>
        public const int UnknownQuadrantRank = 9;

        /// <summary>
        /// The order the follow-up list is shown in until the user picks another one, and what
        /// a stored preference we cannot read falls back to. Newest first is what the page has
        /// always looked like, so nothing moves for anyone who never touches the control.
        /// </summary>
        public const TodoSortMode DefaultMode = TodoSortMode.NewestFirst;

        private static readonly Quadrant[] RankedQuadrants =
        {
            Quadrant.Q1UrgentImportant,
            Quadrant.Q2UrgentNotImportant,
            Quadrant.Q3ImportantNotUrgent,
            Quadrant.Q4Neither
        };

        public static TodoSortMode ParseMode(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return DefaultMode;
            }

            foreach (TodoSortMode mode in Enum.GetValues(typeof(TodoSortMode)))
            {
                if (string.Equals(mode.ToString(), value.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return mode;
                }
            }

            return DefaultMode;
        }

        /// <summary>
        /// The instant an item started waiting: the source mail's arrival when we know it,
        /// otherwise when the item was created.
        /// </summary>
        public static DateTime WaitingSince(TodoItem item)
        {
            if (item == null)
            {
                return DateTime.UtcNow;
            }

            return item.SourceReceivedOn ?? item.CreatedOn;
        }

        /// <summary>
        /// Maps a stored quadrant hint back to its rank. Hints hold the localized name the
        /// user saw when the item was created (「緊急重要」 / "Urgent + Important"), so the
        /// match is done against the current language's names — an item created in another
        /// language falls back to the unknown rank instead of throwing.
        /// </summary>
        public static int QuadrantRank(string hint)
        {
            if (string.IsNullOrEmpty(hint))
            {
                return UnknownQuadrantRank;
            }

            var trimmed = hint.Trim();
            for (var i = 0; i < RankedQuadrants.Length; i++)
            {
                if (string.Equals(trimmed, Strings.QuadrantName(RankedQuadrants[i]), StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return UnknownQuadrantRank;
        }

        /// <summary>
        /// Returns a new list in display order. Finished items always sort last — the page
        /// is about what still needs an answer — and the sort is stable, so items that
        /// compare equal keep the order the store returned.
        /// </summary>
        public static List<TodoItem> Sort(IEnumerable<TodoItem> items, TodoSortMode mode, DateTime nowUtc)
        {
            var list = (items ?? Enumerable.Empty<TodoItem>()).Where(i => i != null).ToList();
            if (list.Count <= 1)
            {
                return list;
            }

            IOrderedEnumerable<TodoItem> ordered = list.OrderBy(i => i.Status == TodoStatus.Done);
            switch (mode)
            {
                case TodoSortMode.NewestFirst:
                    // Newest arrival first, which is what the stored order already produced;
                    // naming it lets the page offer it and lets the choice survive a restart.
                    ordered = ordered.ThenByDescending(i => WaitingSince(i));
                    break;
                case TodoSortMode.LongestWaiting:
                    ordered = ordered.ThenBy(i => WaitingSince(i));
                    break;
                case TodoSortMode.Quadrant:
                    ordered = ordered
                        .ThenBy(i => QuadrantRank(i.QuadrantHint))
                        .ThenBy(i => WaitingSince(i));
                    break;
                case TodoSortMode.RecentlyUpdated:
                    ordered = ordered.ThenByDescending(i => i.UpdatedOn);
                    break;
                case TodoSortMode.DueSoonest:
                    // Undated items must not crowd the top of a "what is due next" list,
                    // so they are ranked behind everything that has a deadline.
                    ordered = ordered
                        .ThenBy(i => i.DueOn.HasValue ? 0 : 1)
                        .ThenBy(i => i.DueOn ?? DateTime.MaxValue);
                    break;
                default:
                    // Manual: keep the stored order within each status group.
                    break;
            }

            return ordered.ToList();
        }
    }
}
