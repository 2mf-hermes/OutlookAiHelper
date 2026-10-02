using System;
using System.Collections.Generic;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// Decides which Outlook folders a scan must skip.
    /// Deleted mail is out of scope: the Deleted Items folder of every store
    /// (primary mailbox + archive PST) and anything filed underneath it.
    /// </summary>
    public static class ScanFolderFilter
    {
        /// <summary>Outlook OlDefaultFolders.olFolderDeletedItems.</summary>
        public const int OlFolderDeletedItems = 3;

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

        /// <summary>
        /// True when the folder must be neither scanned nor walked into.
        /// Matches on the Deleted Items name, or on the EntryID of a store's
        /// default Deleted Items folder (this survives a renamed folder).
        /// </summary>
        public static bool IsExcluded(string folderName, string entryId, ICollection<string> excludedEntryIds)
        {
            if (IsDeletedFolderName(folderName))
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
