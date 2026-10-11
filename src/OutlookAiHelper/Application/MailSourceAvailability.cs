using System;
using System.Collections.Generic;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Application
{
    /// <summary>
    /// The rules the settings page needs in order to draw a source list, kept away from
    /// the UI and from the reader factory so they can be tested on their own: what the
    /// list contains, when a source may be removed, when there is anything new to add,
    /// and what stays active after a removal.
    ///
    /// Every rule here has the same bias: never leave the user without a readable
    /// mailbox. The local Outlook source is always on the list and is never removable,
    /// because it is the fallback every degraded path in MailSourceRuntime lands on.
    /// </summary>
    public static class MailSourceAvailability
    {
        /// <summary>
        /// The cloud kinds this build knows how to name, in the order they are offered.
        /// The local source is deliberately absent: it is never something to "add".
        /// </summary>
        public static readonly MailSourceKind[] CloudKinds = { MailSourceKind.Graph };

        /// <summary>
        /// The rows to show: the local source first, then every configured source this
        /// build can name, in the order the settings file holds them. The local source
        /// appears even when the settings file never stored one, and a stored local
        /// source is not listed twice.
        /// </summary>
        public static IList<MailSourceProfile> DisplayList(IList<MailSourceProfile> configured)
        {
            var shown = new List<MailSourceProfile>();
            shown.Add(Local(configured));

            if (configured != null)
            {
                for (var i = 0; i < configured.Count; i++)
                {
                    var source = configured[i];
                    if (source == null || string.IsNullOrEmpty(source.Id))
                    {
                        continue;
                    }

                    if (source.Kind == MailSourceKind.LocalOutlook)
                    {
                        continue; // already listed above, in its stored form
                    }

                    if (Contains(shown, source.Id))
                    {
                        continue;
                    }

                    shown.Add(source);
                }
            }

            return shown;
        }

        /// <summary>
        /// A source may be removed unless it is the local one. The local source is the
        /// floor of the whole feature: remove it and the app has nothing to read.
        /// </summary>
        public static bool IsRemovable(MailSourceProfile source)
        {
            return source != null && source.Kind != MailSourceKind.LocalOutlook;
        }

        /// <summary>
        /// The kinds in <paramref name="ordered"/> that this build can actually open.
        /// Empty is the normal state today: a kind can be named in the settings file
        /// long before a reader exists for it (see MailReaderFactory.IsImplemented).
        /// </summary>
        public static IList<MailSourceKind> OfferableKinds(Func<MailSourceKind, bool> isImplemented, IList<MailSourceKind> ordered)
        {
            var usable = new List<MailSourceKind>();
            if (isImplemented == null || ordered == null)
            {
                return usable;
            }

            for (var i = 0; i < ordered.Count; i++)
            {
                if (isImplemented(ordered[i]))
                {
                    usable.Add(ordered[i]);
                }
            }

            return usable;
        }

        /// <summary>
        /// The first kind the settings page would add, or null when there is none. With a
        /// single cloud kind there is nothing to choose between, so the add button can
        /// use this instead of opening a picker.
        /// </summary>
        public static MailSourceKind? FirstOfferableKind(Func<MailSourceKind, bool> isImplemented, IList<MailSourceKind> ordered)
        {
            var usable = OfferableKinds(isImplemented, ordered);
            if (usable.Count == 0)
            {
                return null;
            }

            return usable[0];
        }

        /// <summary>
        /// True when this build can offer any source other than the local one, which is
        /// what makes an "add source" control worth showing. False hides it: a button
        /// that could only ever fail is worse than no button.
        /// </summary>
        public static bool CanOfferNew(Func<MailSourceKind, bool> isImplemented)
        {
            return FirstOfferableKind(isImplemented, CloudKinds) != null;
        }

        /// <summary>
        /// True when the first-run step still owes the user a choice. A build that can
        /// read only the local mailbox has nothing to ask, so the sheet never opens even
        /// on a brand new install; the question is recorded as unasked and will be put
        /// the first time this build can offer a second source.
        /// </summary>
        public static bool NeedsFirstRunChoice(Func<MailSourceKind, bool> isImplemented, bool alreadyAsked)
        {
            return !alreadyAsked && CanOfferNew(isImplemented);
        }

        /// <summary>
        /// Which source id is active after removing one. Removing the active source hands
        /// the app back to the local mailbox rather than to nothing, matching what
        /// MailSourceRuntime does with an id that no longer resolves.
        /// </summary>
        public static string ActiveIdAfterRemoval(string activeId, string removedId)
        {
            if (!string.IsNullOrEmpty(activeId) && activeId != removedId)
            {
                return activeId;
            }

            return MailSourceProfile.LocalId;
        }

        private static MailSourceProfile Local(IList<MailSourceProfile> configured)
        {
            if (configured != null)
            {
                for (var i = 0; i < configured.Count; i++)
                {
                    var source = configured[i];
                    if (source != null
                        && source.Kind == MailSourceKind.LocalOutlook
                        && !string.IsNullOrEmpty(source.Id))
                    {
                        return source;
                    }
                }
            }

            return MailSourceProfile.Local();
        }

        private static bool Contains(IList<MailSourceProfile> list, string id)
        {
            for (var i = 0; i < list.Count; i++)
            {
                if (list[i] != null && list[i].Id == id)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
