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
            var item = TodoItem.Create(title, mail.EntryId, quadrantHint, SenderOf(mail), ToUtc(mail.ReceivedOn));
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

        public void ToggleDone(string id)
        {
            var item = Find(id);
            if (item == null)
            {
                return;
            }

            item.Status = item.Status == TodoStatus.Done ? TodoStatus.Open : TodoStatus.Done;
            item.UpdatedOn = DateTime.UtcNow;
            Persist();
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

        /// <summary>
        /// Converts Outlook's ReceivedTime to a UTC instant. Outlook reports local time
        /// with an unspecified Kind, so an unspecified value is read as local — the same
        /// interpretation the mail list itself uses when it prints the timestamp.
        /// </summary>
        private static DateTime? ToUtc(DateTime value)
        {
            if (value == default(DateTime))
            {
                return null;
            }

            return value.Kind == DateTimeKind.Utc ? value : value.ToUniversalTime();
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
