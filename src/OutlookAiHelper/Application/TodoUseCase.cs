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
            var item = TodoItem.Create(title, mail.EntryId, quadrantHint);
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
