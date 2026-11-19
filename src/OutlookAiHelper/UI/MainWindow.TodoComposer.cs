using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using OutlookAiHelper.Adapters.Logging;
using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// Wiring between the main window and <see cref="TodoComposerSheet"/>: opening the
    /// sheet from the mail panel and the detail panel, and turning the sheet's draft into
    /// a real follow-up (file copies first, then the store).
    ///
    /// The split matters: the sheet never touches storage, and a file that fails to copy
    /// still leaves a saved follow-up plus a count the user can see, instead of losing
    /// everything they typed.
    /// </summary>
    public partial class MainWindow
    {
        /// <summary>Where copied attachments are kept, outside the repository.</summary>
        private readonly TodoAttachmentStore _attachmentStore = new TodoAttachmentStore();

        private Button _todoEditButton;

        /// <summary>Opens the sheet for a new follow-up linked to the selected mail.</summary>
        private void OpenTodoComposerForMail()
        {
            if (_selected == null || _selected.Mail == null)
            {
                return;
            }

            OpenTodoComposer(null, _selected.Mail);
        }

        /// <summary>
        /// Opens the sheet for a new follow-up, with no mail attached when one is not given.
        /// The quadrant is taken from the selected mail because a manual follow-up has no
        /// classification of its own, and an empty quadrant leaves it in no box at all.
        /// </summary>
        private void OpenTodoComposer(string prefillTitle, MailSummary mail)
        {
            var draft = TodoComposerSheet.Show(
                this,
                mail,
                null,
                null,
                KnownTodoCategories(),
                prefillTitle);

            if (draft != null)
            {
                var quadrant = _selected != null && _selected.Classification != null
                    ? Strings.QuadrantName(_selected.Classification.Quadrant)
                    : null;

                ApplyTodoDraft(null, draft, quadrant);
            }
        }

        /// <summary>Opens the sheet for a follow-up typed in by hand, with nothing attached.</summary>
        private void OpenTodoComposer(string prefillTitle)
        {
            OpenTodoComposer(prefillTitle, null);
        }

        /// <summary>Opens the sheet to edit an existing follow-up.</summary>
        private void OpenTodoComposerForEdit(TodoItem item)
        {
            if (item == null)
            {
                return;
            }

            var found = FindScannedMail(item.SourceEntryId);
            var source = found == null ? null : found.Mail;

            var draft = TodoComposerSheet.Show(
                this,
                source,
                item,
                item.Attachments,
                KnownTodoCategories());

            if (draft != null)
            {
                ApplyTodoDraft(item, draft, item.QuadrantHint);
            }
        }

        /// <summary>
        /// Turns a draft into a follow-up: copies the chosen files, drops the files the
        /// sheet removed, then persists through the use case so the store stays the single
        /// place items are written.
        /// </summary>
        private void ApplyTodoDraft(TodoItem editing, TodoDraft draft, string quadrantHint)
        {
            if (draft == null)
            {
                return;
            }

            try
            {
                if (editing == null)
                {
                    var item = _todos.AddDraft(draft, quadrantHint);
                    if (item == null)
                    {
                        _statusText.Text = Strings.T("todo.compose.failed");
                        return;
                    }

                    IList<string> errors;
                    var added = _attachmentStore.Import(item.Id, draft.NewFiles, out errors);
                    if (added.Count > 0)
                    {
                        item.Attachments = new List<TodoAttachment>(added);
                        _todos.UpdateDraft(item.Id, draft, item.Attachments);
                    }

                    LogAttachmentErrors(errors);
                    RefreshTodoEntryIds();
                    _mailBadgeStale = true;
                    BindTodos();
                    _statusText.Text = errors.Count == 0
                        ? Strings.T("todo.added")
                        : string.Format(Strings.T("todo.attachments.failed"), errors.Count);
                    ShowPage("todo");
                    return;
                }

                // Edit: keep what is still there, delete what the sheet dropped, copy the new
                // files, then hand the resolved list to the use case in one write.
                var kept = new List<TodoAttachment>();
                if (editing.Attachments != null)
                {
                    foreach (var attachment in editing.Attachments)
                    {
                        if (attachment != null && !draft.RemovedStoredNames.Contains(attachment.StoredName))
                        {
                            kept.Add(attachment);
                        }
                    }
                }

                foreach (var storedName in draft.RemovedStoredNames)
                {
                    _attachmentStore.Remove(editing.Id, storedName);
                }

                IList<string> importErrors;
                var newFiles = _attachmentStore.Import(editing.Id, draft.NewFiles, out importErrors);
                foreach (var created in newFiles)
                {
                    kept.Add(created);
                }

                _todos.UpdateDraft(editing.Id, draft, kept);
                LogAttachmentErrors(importErrors);
                RefreshTodoEntryIds();
                _mailBadgeStale = true;
                BindTodos();

                var refreshed = _todos.Get(editing.Id);
                if (refreshed != null && _detailTodoId == refreshed.Id)
                {
                    ShowTodoDetail(refreshed);
                }

                _statusText.Text = importErrors.Count == 0
                    ? Strings.T("todo.updated")
                    : string.Format(Strings.T("todo.attachments.failed"), importErrors.Count);
            }
            catch (Exception ex)
            {
                FileLogger.Error("ApplyTodoDraft", ex);
                _statusText.Text = Strings.T("todo.compose.failed");
            }
        }

        /// <summary>
        /// The categories already in use, so the editable box suggests them instead of
        /// making the user retype a name and accidentally create a near-duplicate.
        /// </summary>
        private IList<string> KnownTodoCategories()
        {
            var names = new List<string>();
            if (_todos == null || _todos.Items == null)
            {
                return names;
            }

            foreach (var item in _todos.Items)
            {
                if (item == null || string.IsNullOrEmpty(item.Category))
                {
                    continue;
                }

                var known = false;
                foreach (var name in names)
                {
                    if (string.Equals(name, item.Category, StringComparison.CurrentCultureIgnoreCase))
                    {
                        known = true;
                        break;
                    }
                }

                if (!known)
                {
                    names.Add(item.Category);
                }
            }

            names.Sort(StringComparer.CurrentCultureIgnoreCase);
            return names;
        }

        private static void LogAttachmentErrors(IList<string> errors)
        {
            if (errors == null)
            {
                return;
            }

            foreach (var path in errors)
            {
                FileLogger.Info("todo.attachment.failed " + (path ?? string.Empty));
            }
        }

        /// <summary>
        /// Says where a repeating follow-up went after it was ticked off. Null means the
        /// item simply closed, which needs no message.
        /// </summary>
        private void AnnounceRepeat(DateTime? next)
        {
            if (!next.HasValue || _statusText == null)
            {
                return;
            }

            var day = TodoDates.LocalDayOf(next.Value);
            _statusText.Text = string.Format(
                CultureInfo.CurrentCulture,
                Strings.T("todo.repeat.next"),
                day.ToString("d", CultureInfo.CurrentCulture));
        }
    }
}
