using System;
using System.Collections.Generic;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// Decides which Outlook folders a scan must skip.
    /// Out of scope: the Deleted Items folder of every store (primary mailbox +
    /// archive PST) and anything filed underneath it, plus Junk E-mail — junk is
    /// triage noise for a "what still needs an answer" list, and mail parked there
    /// is one click away from being marked Not Junk.
    /// </summary>
    public static class ScanFolderFilter
    {
        /// <summary>Outlook OlDefaultFolders.olFolderDeletedItems.</summary>
        public const int OlFolderDeletedItems = 3;

        /// <summary>Outlook OlDefaultFolders.olFolderJunk.</summary>
        public const int OlFolderJunk = 23;

        /// <summary>
        /// Deleted Items folder name in the Outlook UI languages we ship. Used as a
        /// secondary signal when a store's default folder cannot be resolved by type
        /// (renamed folders, odd or partially loaded profiles).
        /// </summary>
        private static readonly string[] DeletedFolderNames =
        {
            "Deleted Items",
            "刪除的郵件",
            "已删除邮件",
            "已删除的邮件"
        };

        /// <summary>
        /// Junk folder name in the Outlook UI languages we ship. Outlook spells this
        /// differently per locale ("Junk E-mail" / "垃圾郵件"), so the same
        /// EntryID-by-type signal is the primary detector and this is the fallback.
        /// </summary>
        private static readonly string[] JunkFolderNames =
        {
            "Junk E-mail",
            "Junk Email",
            "垃圾郵件",
            "垃圾邮件",
            "垃圾電郵"
        };

        public static bool IsDeletedFolderName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            var trimmed = name.Trim();
            foreach (var candidate in DeletedFolderNames)
            {
                if (string.Equals(trimmed, candidate, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool IsJunkFolderName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            var trimmed = name.Trim();
            foreach (var candidate in JunkFolderNames)
            {
                if (string.Equals(trimmed, candidate, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// True when the folder name alone marks it out of scope (deleted or junk).
        /// A mailbox may also hold user folders the user chose to name this way —
        /// excluding those is the same outcome the user asked for, so no extra guard.
        /// </summary>
        public static bool IsExcludedFolderName(string name)
        {
            return IsDeletedFolderName(name) || IsJunkFolderName(name);
        }

        /// <summary>
        /// True when the folder must be neither scanned nor walked into.
        /// Matches on the excluded names, or on the EntryID of a store's default
        /// Deleted Items / Junk E-mail folder (this survives a renamed folder).
        /// </summary>
        public static bool IsExcluded(string folderName, string entryId, ICollection<string> excludedEntryIds)
        {
            if (IsExcludedFolderName(folderName))
            {
                return true;
            }

            if (string.IsNullOrEmpty(entryId) || excludedEntryIds == null)
            {
                return false;
            }

            foreach (var id in excludedEntryIds)
            {
                if (string.Equals(id, entryId, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
