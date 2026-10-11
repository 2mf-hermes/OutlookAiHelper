using System;
using System.Runtime.Serialization;

namespace OutlookAiHelper.Adapters.Mail
{
    /// <summary>
    /// The parts of an answer that are about the page rather than about a mail: where the next page
    /// is, and how much is unread.
    ///
    /// Kept apart from GraphJson on purpose. That class holds the documents whose fields are asked
    /// for one by one and pinned by a test that refuses any field nobody reads; these two values are
    /// handed back by the service without being asked for, so they would weaken that guard if they
    /// lived beside it. Both still go through the shared json object check, because a page that is
    /// not a json object is a failed read here as much as it is there.
    /// </summary>
    internal static class GraphPage
    {
        /// <summary>
        /// The url of the next page, or null when this was the last one.
        ///
        /// The service only sends a next link when there is more to read, so null is the ordinary
        /// case rather than a failure -- which is why the reader stops paging on null and never
        /// treats it as a short read.
        /// </summary>
        internal static string NextLink(string json)
        {
            var page = JsonBody.Read<PageDocument>(json);
            if (page == null || string.IsNullOrEmpty(page.NextLink))
            {
                return null;
            }

            var link = page.NextLink.Trim();
            return link.Length == 0 ? null : link;
        }

        /// <summary>
        /// A folder's unread total, or null when the answer does not carry one.
        ///
        /// Null rather than zero, because zero is a fact about the mailbox and a missing count is
        /// not: the probe folds this into a signature, and a failed read must not look like an
        /// empty inbox.
        /// </summary>
        internal static int? UnreadCount(string json)
        {
            var folder = JsonBody.Read<FolderCountDocument>(json);
            if (folder == null)
            {
                return null;
            }

            return folder.UnreadItemCount;
        }

        /// <summary>One page of a collection: only the next link, which is all paging needs.</summary>
        [DataContract]
        private sealed class PageDocument
        {
            [DataMember(Name = "@odata.nextLink")]
            public string NextLink { get; set; }
        }

        /// <summary>A folder asked for one field: its unread total.</summary>
        [DataContract]
        private sealed class FolderCountDocument
        {
            [DataMember(Name = "unreadItemCount")]
            public int? UnreadItemCount { get; set; }
        }
    }
}
