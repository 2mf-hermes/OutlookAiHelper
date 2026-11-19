using System;
using System.Collections.Generic;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Application
{
    public sealed class TodoUseCase
    {
        private readonly ITodoStore _store;
        private readonly List<TodoItem> _items;

        public TodoUseCase(ITodoStore store)
        {
            _store = store;
            _items = new List<TodoItem>(_store.Load() ?? new List<TodoItem>());
        }

        public IList<TodoItem> Items
        {
            get { return _items; }
        }

        public void Reload()
        {
            _items.Clear();
            var loaded = _store.Load() ?? new List<TodoItem>();
            _items.AddRange(loaded);
        }

        public void ClearAll()
        {
            _items.Clear();
            Persist();
        }

        public TodoItem AddFromMail(MailSummary mail, string quadrantHint)
        {
            if (mail == null)
            {
                return null;
            }

            var title = string.IsNullOrEmpty(mail.Subject) ? (mail.FromName ?? "Untitled") : mail.Subject;
            var item = TodoItem.Create(title, mail.EntryId, quadrantHint, SenderOf(mail), TodoDates.ToUtc(mail.ReceivedOn));
            _items.Insert(0, item);
            Persist();
            return item;
        }

        public TodoItem AddManual(string title)
        {
            if (string.IsNullOrEmpty(title))
            {
                return null;
            }

            var item = TodoItem.Create(title.Trim());
            _items.Insert(0, item);
            Persist();
            return item;
        }

        /// <summary>
        /// Adds what the composer collected. The item is inserted at the top like the
        /// other add paths, so a follow-up just created is where the user looks first.
        /// </summary>
        public TodoItem AddDraft(TodoDraft draft, string quadrantHint)
        {
            if (draft == null || string.IsNullOrWhiteSpace(draft.Title))
            {
                return null;
            }

            var item = TodoItem.Create(
                draft.Title.Trim(),
                draft.SourceEntryId,
                quadrantHint,
                draft.SourceFrom,
                draft.SourceReceivedOn);
            ApplyDraft(item, draft);
            _items.Insert(0, item);
            Persist();
            return item;
        }

        /// <summary>
        /// Copies an edited draft onto an existing item. Attachments are passed in
        /// already resolved (the caller copied the new files and dropped the removed
        /// ones), which keeps file I/O out of the use case. A null list means "leave
        /// the attachments alone".
        /// </summary>
        public void UpdateDraft(string id, TodoDraft draft, IList<TodoAttachment> attachments = null)
        {
            var item = Find(id);
            if (item == null || draft == null)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(draft.Title))
            {
                item.Title = draft.Title.Trim();
            }

            ApplyDraft(item, draft);
            if (attachments != null)
            {
                item.Attachments = new List<TodoAttachment>(attachments);
            }

            item.UpdatedOn = DateTime.UtcNow;
            Persist();
        }

        public TodoItem Get(string id)
        {
            return Find(id);
        }

        /// <summary>
        /// Flips the item. Completing a repeating follow-up does not leave a finished
        /// row behind: the due date moves to the next occurrence (and the reminder by
        /// the same distance) and the item stays open, which is what the mainstream
        /// task apps do. The next due instant is returned so the caller can say when,
        /// and null means an ordinary open/done flip.
        /// </summary>
        public DateTime? ToggleDone(string id)
        {
            var item = Find(id);
            if (item == null)
            {
                return null;
            }

            if (item.Status == TodoStatus.Open && TodoRecurrence.IsRepeating(item.Repeat) && item.DueOn.HasValue)
            {
                var anchor = item.RepeatAnchorDay > 0
                    ? item.RepeatAnchorDay
                    : item.DueOn.Value.ToLocalTime().Day;
                var next = TodoRecurrence.Next(item.DueOn.Value, item.Repeat, anchor);
                var shift = next - item.DueOn.Value;
                item.DueOn = next;
                if (item.ReminderOn.HasValue)
                {
                    item.ReminderOn = item.ReminderOn.Value + shift;
                }

                item.RepeatAnchorDay = anchor;
                item.UpdatedOn = DateTime.UtcNow;
                Persist();
                return next;
            }

            item.Status = item.Status == TodoStatus.Done ? TodoStatus.Open : TodoStatus.Done;
            item.UpdatedOn = DateTime.UtcNow;
            Persist();
            return null;
        }

        public void SetNote(string id, string note)
        {
            var item = Find(id);
            if (item == null)
            {
                return;
            }

            item.Note = note ?? string.Empty;
            item.UpdatedOn = DateTime.UtcNow;
            Persist();
        }

        /// <summary>
        /// Sets the due instant, or clears it when <paramref name="dueOn"/> is null.
        /// The caller passes UTC (the UI converts the picked local day's end), so the
        /// overdue test is a plain comparison against DateTime.UtcNow.
        /// </summary>
        public void SetDue(string id, DateTime? dueOn)
        {
            var item = Find(id);
            if (item == null)
            {
                return;
            }

            item.DueOn = dueOn;
            item.UpdatedOn = DateTime.UtcNow;
            Persist();
        }

        /// <summary>
        /// Drops every finished follow-up in one pass. Returns how many were removed so
        /// the caller can report "已清除 N 筆" instead of silently emptying the list.
        /// </summary>
        public int ClearDone()
        {
            var removed = _items.RemoveAll(i => i != null && i.Status == TodoStatus.Done);
            if (removed > 0)
            {
                Persist();
            }

            return removed;
        }

        private static void ApplyDraft(TodoItem item, TodoDraft draft)
        {
            item.Note = draft.Note ?? string.Empty;
            item.Category = (draft.Category ?? string.Empty).Trim();
            item.DueOn = draft.DueOn;
            item.ReminderOn = draft.ReminderOn;
            item.Repeat = draft.Repeat;
            item.RepeatAnchorDay = draft.RepeatAnchorDay;
            item.SourceEntryId = draft.SourceEntryId;
            item.SourceFrom = draft.SourceFrom;
            item.SourceReceivedOn = draft.SourceReceivedOn;
        }

        public void Remove(string id)
        {
            var item = Find(id);
            if (item == null)
            {
                return;
            }

            _items.Remove(item);
            Persist();
        }

        private static string SenderOf(MailSummary mail)
        {
            if (mail == null)
            {
                return null;
            }

            if (!string.IsNullOrEmpty(mail.FromName))
            {
                return mail.FromName;
            }

            return mail.FromAddress;
        }

        private TodoItem Find(string id)
        {
            if (string.IsNullOrEmpty(id))
            {
                return null;
            }

            foreach (var item in _items)
            {
                if (string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase))
                {
                    return item;
                }
            }

            return null;
        }

        private void Persist()
        {
            _store.Save(_items);
        }
    }
}
