using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
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

        /// <summary>Follow-up whose related mails the block is showing (null when none).</summary>
        private string _relatedTodoId;

        /// <summary>
        /// The related-mails block and the list inside it. Built once with the detail panel and
        /// filled per selection; both are null until BuildDetail has run.
        /// </summary>
        private StackPanel _relatedBlock;
        private StackPanel _relatedPanel;

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
                    RenderRelatedMails(item, mail);
                    HighlightTodoRows();
                    return;
                }

                ShowTodoOnlyDetail(item);
                RenderTodoActions(item);
                RenderRelatedMails(item, null);
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

        /// <summary>
        /// The block that lists every mail of the follow-up's topic, newest first by default.
        /// It is built with the panel and shown or hidden per selection, so the panel keeps one
        /// layout instead of growing a second heading.
        /// </summary>
        private StackPanel BuildRelatedBlock()
        {
            var block = new StackPanel { Visibility = Visibility.Collapsed };
            block.Children.Add(UiKit.Subtitle(Strings.T("todo.detail.related")));

            _relatedPanel = new StackPanel();
            block.Children.Add(_relatedPanel);
            return block;
        }

        /// <summary>
        /// Fills the related-mails block for a follow-up. One related mail would only repeat the
        /// mail already shown above, so the block stays down unless the topic really holds more.
        /// <paramref name="source"/> is the scanned mail behind the follow-up when there is one,
        /// and it is also what defines the topic.
        /// </summary>
        private void RenderRelatedMails(TodoItem item, ScanResultItem source)
        {
            if (_relatedBlock == null || _relatedPanel == null || item == null)
            {
                return;
            }

            _relatedTodoId = item.Id;
            _relatedPanel.Children.Clear();

            var related = RelatedMails(item, source);
            if (related.Count <= 1)
            {
                _relatedBlock.Visibility = Visibility.Collapsed;
                return;
            }

            for (var i = 0; i < related.Count; i++)
            {
                _relatedPanel.Children.Add(BuildRelatedMailRow(related[i]));
            }

            _relatedBlock.Visibility = Visibility.Visible;
        }

        /// <summary>Shows or hides the block; the panel itself decides when it has content.</summary>
        private void SetRelatedMailsVisible(bool visible)
        {
            if (_relatedBlock != null)
            {
                _relatedBlock.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        /// <summary>
        /// Re-orders an already open block after the sort setting changed. Nothing to do when no
        /// follow-up is on screen.
        /// </summary>
        private void RefreshRelatedMails()
        {
            if (_relatedBlock == null
                || _relatedBlock.Visibility != Visibility.Visible
                || string.IsNullOrEmpty(_relatedTodoId))
            {
                return;
            }

            var item = FindTodo(_relatedTodoId);
            if (item == null)
            {
                return;
            }

            RenderRelatedMails(item, FindScannedMail(item.SourceEntryId));
        }

        /// <summary>
        /// Every scanned mail in the follow-up's topic. The topic comes from the follow-up's own
        /// mail while the scan still holds it, and from the subject it stored otherwise, so the
        /// block survives the mail dropping out of the scan window.
        /// </summary>
        private List<ScanResultItem> RelatedMails(TodoItem item, ScanResultItem source)
        {
            var key = source != null && source.Mail != null
                ? MailThreads.KeyOf(source.Mail)
                : MailThreads.KeyOfSubject(item.Title);

            if (string.IsNullOrEmpty(key))
            {
                return new List<ScanResultItem>();
            }

            var related = _items.Where(i =>
                i != null
                && i.Mail != null
                && string.Equals(MailThreads.KeyOf(i.Mail), key, StringComparison.OrdinalIgnoreCase));

            return MailOrdering.SortMails(related, CurrentMailSort());
        }

        /// <summary>One related mail: subject, sender and time, with its quadrant at the end.</summary>
        private FrameworkElement BuildRelatedMailRow(ScanResultItem mail)
        {
            var titles = new StackPanel();
            titles.Children.Add(new TextBlock
            {
                Text = mail.Mail.Subject,
                FontSize = UiKit.TypeSubhead,
                Foreground = Theme.InkBrush,
                TextTrimming = TextTrimming.CharacterEllipsis
            });
            titles.Children.Add(new TextBlock
            {
                Text = mail.Mail.FromName + " \u00b7 " + mail.Mail.ReceivedOn.ToString("MM/dd HH:mm"),
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.MutedBrush,
                Margin = new Thickness(0, 2, 0, 0),
                TextTrimming = TextTrimming.CharacterEllipsis
            });

            var grid = (System.Windows.Controls.Grid)UiKit.ListRow(
                null,
                titles,
                new TextBlock
                {
                    Text = Strings.QuadrantName(mail.Classification.Quadrant),
                    FontSize = UiKit.TypeCaption,
                    FontWeight = FontWeights.Medium,
                    Foreground = Theme.SecondaryBrush,
                    VerticalAlignment = VerticalAlignment.Center
                });
            grid.Tag = mail;
            grid.Cursor = Cursors.Hand;
            grid.Background = Brushes.Transparent;
            AutomationProperties.SetName(
                grid,
                string.Format(Strings.T("mail.thread.member"), mail.Mail.Subject));
            grid.MouseLeftButtonUp += (s, e) => OnRelatedMailClick(mail);
            return grid;
        }

        /// <summary>
        /// Shows one of the follow-up's related mails in the same panel. The plain-mail selection
        /// underneath switches the follow-up block off, so it is put straight back.
        /// </summary>
        private void OnRelatedMailClick(ScanResultItem mail)
        {
            var item = FindTodo(_relatedTodoId);
            if (item == null || mail == null)
            {
                return;
            }

            try
            {
                _lastPicked = mail;
                OnMailSelectedDetail(mail);
                RenderTodoActions(item);
                RenderRelatedMails(item, mail);
                HighlightTodoRows();
            }
            catch (Exception ex)
            {
                FileLogger.Error("OnRelatedMailClick", ex);
            }
        }

        /// <summary>Follow-up by id, for the panel's own lookups (null once it is gone).</summary>
        private TodoItem FindTodo(string id)
        {
            if (string.IsNullOrEmpty(id) || _todos == null || _todos.Items == null)
            {
                return null;
            }

            return _todos.Items.FirstOrDefault(t =>
                t != null && string.Equals(t.Id, id, StringComparison.OrdinalIgnoreCase));
        }
    }
}
