using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using OutlookAiHelper.Adapters.Logging;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// Follow-up (待回待辦) detail surface: clicking a row shows the underlying mail in
    /// the right-hand panel. Kept in its own partial file, like MainWindow.Update.cs, so
    /// MainWindow.cs stays readable.
    ///
    /// When the source mail is still inside the current scan we reuse the existing mail
    /// detail (from/time/folder, classification reasons, actions). When it is not - the
    /// scan window moved on, or the mail was deleted - we still show everything the
    /// follow-up itself recorded, so the panel is never a dead end.
    /// </summary>
    public sealed partial class MainWindow
    {
        /// <summary>Entry id the panel is showing while no scan result backs it (may be null).</summary>
        private string _detailEntryId;

        /// <summary>Follow-up whose row the panel is showing (null when none).</summary>
        private string _detailTodoId;

        private void ShowTodoDetail(TodoItem item)
        {
            try
            {
                if (item == null)
                {
                    return;
                }

                _detailTodoId = item.Id;

                var mail = FindScannedMail(item.SourceEntryId);
                if (mail != null)
                {
                    _detailEntryId = null;
                    OnMailSelectedDetail(mail);
                    RenderTodoActions(item);
                    HighlightTodoRows();
                    return;
                }

                ShowTodoOnlyDetail(item);
                RenderTodoActions(item);
                HighlightTodoRows();
            }
            catch (Exception ex)
            {
                FileLogger.Error("ShowTodoDetail", ex);
            }
        }

        /// <summary>Detail for a follow-up whose source mail is not in the current scan.</summary>
        private void ShowTodoOnlyDetail(TodoItem item)
        {
            _selected = null;
            _detailEntryId = item.SourceEntryId;

            if (_detailHost == null)
            {
                return;
            }

            _detailHost.Visibility = Visibility.Visible;
            UpdatePanelWidths();

            if (_reclassifyButton == null)
            {
                return;
            }

            // Only the actions that need a scanned mail are switched off; opening the
            // original in Outlook works off the stored entry id.
            _reclassifyButton.IsEnabled = false;
            _addTodoButton.IsEnabled = false;
            if (_aiAction != null)
            {
                _aiAction.IsEnabled = false;
            }

            if (_openOutlookButton != null)
            {
                _openOutlookButton.IsEnabled = !string.IsNullOrEmpty(item.SourceEntryId);
            }

            _detailTitle.Text = item.Title;

            var meta = new StringBuilder();
            meta.Append(Strings.T("todo.detail.state"))
                .Append(": ")
                .Append(Strings.T(item.Status == TodoStatus.Done ? "todo.filter.done" : "todo.filter.open"));

            if (!string.IsNullOrEmpty(item.QuadrantHint))
            {
                meta.Append("\n").Append(Strings.T("todo.detail.quadrant")).Append(": ").Append(item.QuadrantHint);
            }

            meta.Append("\n").Append(Strings.T("todo.detail.created")).Append(": ")
                .Append(item.CreatedOn.ToLocalTime().ToString("yyyy-MM-dd HH:mm"));
            meta.Append("\n").Append(Strings.T("todo.detail.updated")).Append(": ")
                .Append(item.UpdatedOn.ToLocalTime().ToString("yyyy-MM-dd HH:mm"));

            if (item.DueOn.HasValue)
            {
                meta.Append("\n").Append(Strings.T("todo.detail.due")).Append(": ")
                    .Append(item.DueOn.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm"));
            }

            _detailMeta.Text = meta.ToString();

            _reasonPanel.Children.Clear();
            _reasonPanel.Children.Add(new TextBlock
            {
                Text = Strings.T("todo.detail.mailMissing"),
                FontSize = UiKit.TypeSubhead,
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap
            });
        }

        private ScanResultItem FindScannedMail(string entryId)
        {
            if (string.IsNullOrEmpty(entryId))
            {
                return null;
            }

            return _items.FirstOrDefault(i =>
                i != null
                && i.Mail != null
                && string.Equals(i.Mail.EntryId, entryId, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>Marks the row the panel is showing (rows are rebuilt by BindTodos).</summary>
        private void HighlightTodoRows()
        {
            try
            {
                if (_todoList == null)
                {
                    return;
                }

                foreach (var entry in _todoList.Items)
                {
                    var row = entry as Grid;
                    if (row == null)
                    {
                        continue;
                    }

                    var todo = row.Tag as TodoItem;
                    var showing = todo != null
                        && !string.IsNullOrEmpty(_detailTodoId)
                        && todo.Id == _detailTodoId;
                    row.Background = showing ? Theme.RowSelectedBrush : Brushes.Transparent;
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("HighlightTodoRows", ex);
            }
        }

        /// <summary>
        /// True when the click came from a control inside the row (check box, delete
        /// button), so the row's own click-to-preview handler stays out of the way.
        /// </summary>
        private static bool IsFromInteractiveChild(DependencyObject source)
        {
            var node = source;
            while (node != null)
            {
                if (node is ButtonBase)
                {
                    return true;
                }

                node = VisualTreeHelper.GetParent(node);
            }

            return false;
        }
    }
}
