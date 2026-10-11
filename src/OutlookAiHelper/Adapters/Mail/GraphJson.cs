using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Adapters.Mail
{
    /// <summary>
    /// Turns the two Graph payloads this app reads -- a page of messages and a list of folders --
    /// into the app's own models, so nothing above this file has to know what Graph's json looks
    /// like.
    ///
    /// The mapping is a mirror of the Outlook adapter's, not a set of choices of its own. The mail
    /// list, the sort and the "flagged" / "meeting" filters all read the same MailSummary fields,
    /// so a cloud mailbox and a desktop mailbox have to describe a mail the same way or the same
    /// mail would look different depending on where it was read from. That is why a cancellation
    /// counts as a meeting request here (the desktop reader counts Outlook's class 54 too) and why
    /// a completed flag counts as flagged (Outlook's flag status is any non-zero value).
    ///
    /// Nothing here opens a socket: every statement below is pinned by offline tests, which is the
    /// only kind of test available to a build that has no app registration yet.
    /// </summary>
    internal static class GraphJson
    {
        /// <summary>
        /// The well-known name Graph gives the Inbox folder. Matched on instead of the display
        /// name, which is translated and so cannot be matched on.
        /// </summary>
        internal const string InboxWellKnownName = "inbox";

        /// <summary>
        /// The json members behind a message, and the sub-members that ride along inside "from"
        /// and "flag". Kept next to the document types below so the request and the reader cannot
        /// drift apart: a test compares these three lists with the document types' member names
        /// and with the fields GraphApp asks for.
        /// </summary>
        internal static readonly string[] MessageReadFields =
        {
            "id", "subject", "from", "receivedDateTime", "isRead", "flag", "importance", "meetingMessageType",
            "emailAddress", "name", "address", "flagStatus"
        };

        /// <summary>
        /// The json members behind a folder.
        /// </summary>
        internal static readonly string[] FolderReadFields = { "id", "displayName", "wellKnownName" };

        /// <summary>
        /// The json members that are structure rather than content: the array of results and the
        /// body of an error.
        /// </summary>
        internal static readonly string[] EnvelopeFields = { "value", "error", "code", "message" };

        /// <summary>
        /// One page of messages as MailSummary rows. Returns null when the payload cannot be read
        /// at all, and an empty list when it is readable but holds no messages: the caller tells
        /// "the answer was unreadable" apart from "the folder is empty".
        /// </summary>
        internal static IList<MailSummary> ParseMessages(string json, string folderDisplayName)
        {
            var page = Deserialize<MessagePageDocument>(json);
            if (page == null)
            {
                return null;
            }

            var messages = new List<MailSummary>();
            if (page.Value == null)
            {
                return messages;
            }

            var folder = string.IsNullOrEmpty(folderDisplayName) ? "Inbox" : folderDisplayName;
            foreach (var message in page.Value)
            {
                if (message == null || string.IsNullOrEmpty(message.Id))
                {
                    // No id means the row could never be linked to a todo or opened again, so it
                    // is dropped rather than shown as a mail that cannot be acted on.
                    continue;
                }

                messages.Add(new MailSummary
                {
                    EntryId = message.Id,
                    FolderPath = folder,
                    Subject = message.Subject ?? string.Empty,
                    FromName = SenderPart(message, true),
                    FromAddress = SenderPart(message, false),
                    ReceivedOn = ParseReceived(message.ReceivedDateTime),
                    // Graph always sends isRead; a payload that omits it counts the mail as
                    // unread. The one rule that reads this field treats unread mail as mail that
                    // still needs an answer, and hiding a mail from that rule is worse than
                    // showing one the user has already read.
                    IsRead = message.IsRead == true,
                    HasFlag = IsFlagged(message.Flag),
                    IsMeetingRequest = IsMeetingRequestType(message.MeetingMessageType),
                    Importance = ParseImportance(message.Importance)
                });
            }

            return messages;
        }

        /// <summary>
        /// The folders Graph lists for the mailbox, in the order it returned them. Entries without
        /// an id or a name are dropped; null means the payload was unreadable.
        /// </summary>
        internal static IList<GraphFolderInfo> ParseFolders(string json)
        {
            var page = Deserialize<FolderPageDocument>(json);
            if (page == null)
            {
                return null;
            }

            var folders = new List<GraphFolderInfo>();
            if (page.Value == null)
            {
                return folders;
            }

            foreach (var folder in page.Value)
            {
                if (folder == null || string.IsNullOrEmpty(folder.Id) || string.IsNullOrEmpty(folder.DisplayName))
                {
                    continue;
                }

                folders.Add(new GraphFolderInfo
                {
                    Id = folder.Id,
                    DisplayName = folder.DisplayName,
                    WellKnownName = folder.WellKnownName ?? string.Empty
                });
            }

            return folders;
        }

        /// <summary>
        /// The folder names the settings page offers, with the Inbox first -- the same shape the
        /// desktop reader returns, so the folder picker does not change when the source does.
        /// </summary>
        internal static IList<string> FolderNames(IList<GraphFolderInfo> folders)
        {
            var ordered = new List<string>();
            if (folders == null)
            {
                return ordered;
            }

            foreach (var folder in folders)
            {
                if (IsInbox(folder))
                {
                    ordered.Add(folder.DisplayName);
                }
            }

            foreach (var folder in folders)
            {
                if (!IsInbox(folder))
                {
                    ordered.Add(folder.DisplayName);
                }
            }

            return ordered;
        }

        /// <summary>
        /// The id of the folder a scan asked for, or null when this mailbox has no such folder.
        /// Matched on the display name first, and on the well-known name for the literal "Inbox",
        /// so a request written in English still finds a mailbox whose Inbox is translated.
        /// </summary>
        internal static string ResolveFolderId(IList<GraphFolderInfo> folders, string displayName)
        {
            if (folders == null || string.IsNullOrEmpty(displayName))
            {
                return null;
            }

            var wanted = displayName.Trim();
            foreach (var folder in folders)
            {
                if (string.Equals(folder.DisplayName, wanted, StringComparison.OrdinalIgnoreCase))
                {
                    return folder.Id;
                }
            }

            if (string.Equals(wanted, "Inbox", StringComparison.OrdinalIgnoreCase))
            {
                foreach (var folder in folders)
                {
                    if (IsInbox(folder))
                    {
                        return folder.Id;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// The "code" of a Graph error body, or null when the text is not one. The code is what a
        /// failure is reported as ("invalid_grant" reads very differently from a network error),
        /// so it is read from the payload rather than guessed from the status code.
        /// </summary>
        internal static string ParseErrorCode(string json)
        {
            var envelope = Deserialize<ErrorDocument>(json);
            if (envelope == null || envelope.Error == null)
            {
                return null;
            }

            return envelope.Error.Code;
        }

        private static bool IsInbox(GraphFolderInfo folder)
        {
            return folder != null
                && string.Equals(folder.WellKnownName, InboxWellKnownName, StringComparison.OrdinalIgnoreCase);
        }

        private static string SenderPart(MessageDocument message, bool name)
        {
            if (message.From == null || message.From.EmailAddress == null)
            {
                return string.Empty;
            }

            var value = name ? message.From.EmailAddress.Name : message.From.EmailAddress.Address;
            return value == null ? string.Empty : value.Trim();
        }

        /// <summary>
        /// Graph stamps a received time in UTC ("2024-05-06T07:08:09Z"); the desktop reader hands
        /// back local time. Everything downstream (the day cutoff, the sort, the time on screen)
        /// compares against the local clock, so the value is converted here. A time that cannot be
        /// read falls back to now, which is what the desktop reader does when Outlook gives none.
        /// </summary>
        private static DateTime ParseReceived(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return DateTime.Now;
            }

            DateTime parsed;
            if (!DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out parsed))
            {
                return DateTime.Now;
            }

            return parsed.Kind == DateTimeKind.Utc ? parsed.ToLocalTime() : parsed;
        }

        /// <summary>
        /// Outlook's flag status is a number whose zero value means "not flagged", so the desktop
        /// reader counts a completed flag as flagged too. Mirrored here.
        /// </summary>
        private static bool IsFlagged(FlagDocument flag)
        {
            if (flag == null || string.IsNullOrEmpty(flag.FlagStatus))
            {
                return false;
            }

            return string.Equals(flag.FlagStatus, "flagged", StringComparison.OrdinalIgnoreCase)
                || string.Equals(flag.FlagStatus, "complete", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Outlook reports a meeting request as message class 53 and a cancellation as 54, and the
        /// desktop reader counts both. Graph's meetingMessageType carries the same distinction.
        /// </summary>
        private static bool IsMeetingRequestType(string meetingMessageType)
        {
            if (string.IsNullOrEmpty(meetingMessageType))
            {
                return false;
            }

            return string.Equals(meetingMessageType, "meetingRequest", StringComparison.OrdinalIgnoreCase)
                || string.Equals(meetingMessageType, "meetingCancelled", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Outlook's importance is 0 for low, 1 for normal and 2 for high. Anything Graph sends
        /// that is not low or high is read as normal, which is also the Outlook default.
        /// </summary>
        private static int ParseImportance(string importance)
        {
            if (string.Equals(importance, "low", StringComparison.OrdinalIgnoreCase))
            {
                return 0;
            }

            if (string.Equals(importance, "high", StringComparison.OrdinalIgnoreCase))
            {
                return 2;
            }

            return 1;
        }

        private static T Deserialize<T>(string json) where T : class
        {
            if (string.IsNullOrEmpty(json))
            {
                return null;
            }

            var trimmed = json.TrimStart();
            if (trimmed.Length == 0 || trimmed[0] != '{')
            {
                // Every body this app reads is a json object. Anything else -- an html sign-in
                // page handed back by a proxy, a bare array, a quoted string -- is a failed read,
                // and saying so is better than reporting an empty mailbox.
                return null;
            }

            try
            {
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    return (T)new DataContractJsonSerializer(typeof(T)).ReadObject(stream);
                }
            }
            catch (Exception)
            {
                // A malformed or truncated payload is a failed read, not a crash: this runs on the
                // UI thread of a mail client.
                return null;
            }
        }

        [DataContract]
        private sealed class MessagePageDocument
        {
            [DataMember(Name = "value")]
            public MessageDocument[] Value { get; set; }
        }

        [DataContract]
        private sealed class MessageDocument
        {
            [DataMember(Name = "id")]
            public string Id { get; set; }

            [DataMember(Name = "subject")]
            public string Subject { get; set; }

            [DataMember(Name = "from")]
            public RecipientDocument From { get; set; }

            [DataMember(Name = "receivedDateTime")]
            public string ReceivedDateTime { get; set; }

            [DataMember(Name = "isRead")]
            public bool? IsRead { get; set; }

            [DataMember(Name = "flag")]
            public FlagDocument Flag { get; set; }

            [DataMember(Name = "importance")]
            public string Importance { get; set; }

            [DataMember(Name = "meetingMessageType")]
            public string MeetingMessageType { get; set; }
        }

        [DataContract]
        private sealed class RecipientDocument
        {
            [DataMember(Name = "emailAddress")]
            public EmailAddressDocument EmailAddress { get; set; }
        }

        [DataContract]
        private sealed class EmailAddressDocument
        {
            [DataMember(Name = "name")]
            public string Name { get; set; }

            [DataMember(Name = "address")]
            public string Address { get; set; }
        }

        [DataContract]
        private sealed class FlagDocument
        {
            [DataMember(Name = "flagStatus")]
            public string FlagStatus { get; set; }
        }

        [DataContract]
        private sealed class FolderPageDocument
        {
            [DataMember(Name = "value")]
            public FolderDocument[] Value { get; set; }
        }

        [DataContract]
        private sealed class FolderDocument
        {
            [DataMember(Name = "id")]
            public string Id { get; set; }

            [DataMember(Name = "displayName")]
            public string DisplayName { get; set; }

            [DataMember(Name = "wellKnownName")]
            public string WellKnownName { get; set; }
        }

        [DataContract]
        private sealed class ErrorDocument
        {
            [DataMember(Name = "error")]
            public ErrorBodyDocument Error { get; set; }
        }

        [DataContract]
        private sealed class ErrorBodyDocument
        {
            [DataMember(Name = "code")]
            public string Code { get; set; }

            [DataMember(Name = "message")]
            public string Message { get; set; }
        }
    }

    /// <summary>
    /// A folder of a cloud mailbox, as much of one as this app uses: the id to read messages
    /// from, the name to show, and the well-known name when Graph calls it a default folder.
    /// </summary>
    internal sealed class GraphFolderInfo
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string WellKnownName { get; set; }
    }
}
