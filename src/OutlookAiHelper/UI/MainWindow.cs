using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Automation;
using OutlookAiHelper.Adapters.Outlook;
using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Abstractions;
using OutlookAiHelper.Core.Classification;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.UI
{
    public sealed partial class MainWindow : Window
    {
        private readonly RuleOptions _ruleOptions = RuleOptions.CreateDefault();
        private readonly OutlookComMailReader _reader = new OutlookComMailReader();
        private readonly JsonOverrideStore _overrideStore = new JsonOverrideStore();
        private readonly JsonSettingsStore _settingsStore = new JsonSettingsStore();
        private readonly PrivacyUseCase _privacy;
        private AppSettings _settings;
        private readonly TodoUseCase _todos;
        private readonly List<ScanResultItem> _items = new List<ScanResultItem>();
        private readonly Dictionary<string, OverrideEntry> _overrides = new Dictionary<string, OverrideEntry>(StringComparer.OrdinalIgnoreCase);

        private ListBox _mailList;
        private ListBox _quadrantFilter;
        private ListBox _todoList;
        private StackPanel _reasonPanel;
        private FrameworkElement _detailHost;
        private ColumnDefinition _leftCol;
        private ColumnDefinition _rightCol;
        private bool _sidebarOpen = true;
        private volatile bool _navBusy;
        private bool _suppressTodoUi;
        private bool _filterUpdating;
        private FrameworkElement _navHost;
        private Button _sidebarToggle;
        private TextBlock _detailTitle;
        private TextBlock _detailMeta;
        private TextBlock _statusText;
        private TextBlock _todoEmpty;
        private ProgressBar _progress;
        private Button _scanButton;
        private Button _cancelButton;
        private Button _reclassifyButton;
        private Button _addTodoButton;
        private Button _navQuadrants;
        private Button _navTodo;
        private Button _navSettings;
        private ComboBox _daysBox;
        private ComboBox _todoFilter;
        private TextBox _todoInput;
        private TextBox _settingsDays;
        private TextBox _settingsUrgent;
        private TextBox _settingsImportant;
        private TextBox _settingsVip;
        private ComboBox _settingsLanguage;
        private CheckBox _aiEnabled;
        private TextBox _aiBaseUrl;
        private TextBox _aiApiKey;
        private ComboBox _aiProviders;
        private ComboBox _aiModel;
        private TextBlock _aiModelHint;
        private bool _aiUiSync;
        private Button _aiAction;
        private Button _openOutlookButton;
        private Panel _emptyPanel;
        private Panel _errorPanel;
        private Panel _listPanel;
        private Panel _quadrantPage;
        private Panel _todoPage;
        private Panel _settingsPage;
        private Quadrant _filter = Quadrant.Q1UrgentImportant;
        private bool _filterAll;
        private string _todoFilterMode = "Open";
        private CancellationTokenSourceLike _cancel;
        private ScanResultItem _selected;
        private ScanResultItem _lastPicked;
        private bool _autoScanDone;

        public MainWindow()
        {
            Title = Strings.T("app.title") + "  ·  v" + AppInfo.Version;
            Width = 1180;
            Height = 760;
            MinWidth = 900;
            MinHeight = 600;
            FontFamily = Theme.UiFont;
            UseLayoutRounding = true;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            try
            {
                TextOptions.SetTextFormattingMode(this, TextFormattingMode.Display);
                TextOptions.SetTextRenderingMode(this, TextRenderingMode.ClearType);
            }
            catch (Exception)
            {
            }

            // acrylic disabled — liquid glass painted with layered brushes

            _settings = _settingsStore.Load();
            Strings.Language = _settings.ToUiLanguage();
            SyncRulesFromSettings(_settings);
            _privacy = new PrivacyUseCase(JsonPaths.RootDirectory);
            _todos = new TodoUseCase(new JsonTodoStore());

            try
            {
                var iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.ico");
                if (File.Exists(iconPath))
                {
                    Icon = new BitmapImage(new Uri(iconPath, UriKind.Absolute));
                }
            }
            catch (Exception)
            {
            }

            EnsureAiProviders();
            Content = BuildChrome();
            ApplySettingsToToolbar();
            RefreshAiProviderCombo();
            ShowEmpty();
            ShowPage("quadrants");
            ConfigureAutoRefresh();
            UpdateTodoBadge();
            Loaded += async (s, e) =>
            {
                if (_autoScanDone)
                {
                    return;
                }

                _autoScanDone = true;
                try
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        _statusText.Text = Strings.T("auto.scan.start");
                    }).Task;

                    await Task.Delay(250);
                    await RunScanAsync();
                }
                catch (Exception ex)
                {
                    ShowError("error.scan.failed");
                    _statusText.Text = ex.Message;
                }
            };
            Loaded += async (s, e) => await StartupUpdateCheckAsync();
        }

        private void SyncRulesFromSettings(AppSettings settings)
        {
            if (settings == null)
            {
                return;
            }

            _ruleOptions.UrgentKeywords = settings.UrgentKeywords != null && settings.UrgentKeywords.Count > 0
                ? new List<string>(settings.UrgentKeywords)
                : RuleOptions.DefaultUrgentKeywords();
            _ruleOptions.ImportantKeywords = settings.ImportantKeywords != null && settings.ImportantKeywords.Count > 0
                ? new List<string>(settings.ImportantKeywords)
                : RuleOptions.DefaultImportantKeywords();
            _ruleOptions.VipAddresses = new List<string>(settings.VipAddresses ?? new List<string>());
        }

        private void ApplySettingsToToolbar()
        {
            if (_settings == null || _daysBox == null)
            {
                return;
            }

            var dayText = _settings.ScanDays.ToString();
            _daysBox.SelectedItem = dayText;
            if (_daysBox.SelectedItem == null)
            {
                _daysBox.Items.Insert(0, dayText);
                _daysBox.SelectedItem = dayText;
            }
        }

        private UIElement BuildChrome()
        {
            // Content backdrop (under glass control layer)
            var root = new Grid();
            root.Children.Add(Theme.AmbientField());

            // Glass control layer floats above content
            var layout = new Grid { Margin = new Thickness(18) };
            _leftCol = new ColumnDefinition { Width = new GridLength(220) };
            _rightCol = new ColumnDefinition { Width = GridLength.Auto };
            layout.ColumnDefinitions.Add(_leftCol);
            layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            layout.ColumnDefinitions.Add(_rightCol);

            var nav = BuildNav();
            _navHost = nav as FrameworkElement;
            _sidebarToggle = CreateSidebarToggle();
            layout.Children.Add(nav);
            Grid.SetColumn(nav, 0);

            var pages = new Grid { Margin = new Thickness(14, 0, 14, 0) };
            _quadrantPage = BuildCenter();
            _todoPage = BuildTodoPage();
            _settingsPage = BuildSettingsPage();
            pages.Children.Add(_quadrantPage);
            pages.Children.Add(_todoPage);
            pages.Children.Add(_settingsPage);
            layout.Children.Add(pages);
            Grid.SetColumn(pages, 1);

            var detail = BuildDetail();
            _detailHost = detail as FrameworkElement;
            if (_detailHost != null)
            {
                _detailHost.Visibility = Visibility.Collapsed;
                _detailHost.Margin = new Thickness(8, 0, 0, 0);
            }
            layout.Children.Add(detail);
            Grid.SetColumn(detail, 2);
            UpdatePanelWidths();

            root.Children.Add(layout);
            return root;
        }

        private void HideDetail()
        {
            if (_detailHost != null)
            {
                _detailHost.Visibility = Visibility.Collapsed;
            }

            _selected = null;
            _detailEntryId = null;
            _detailTodoId = null;
            HighlightTodoRows();
            UpdatePanelWidths();
        }

        private void UpdatePanelWidths()
        {
            if (_leftCol != null)
            {
                // Keep a slim rail when collapsed so the expand button stays clickable.
                _leftCol.Width = _sidebarOpen ? new GridLength(220) : new GridLength(0);
            }

            if (_rightCol != null && _detailHost != null)
            {
                var open = _detailHost.Visibility == Visibility.Visible;
                _rightCol.Width = open ? new GridLength(300) : new GridLength(0);
            }
        }

        private Button CreateSidebarToggle()
        {
            var btn = new Button
            {
                Width = 28,
                Height = 28,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Center,
                Padding = new Thickness(0),
                Content = UiKit.SidebarGlyph(14),
                ToolTip = Strings.T("nav.sidebarToggle")
            };
            // Icon-only: without this a screen reader announces a nameless button, and any
            // UI automation driving the window has nothing to find it by.
            AutomationProperties.SetName(btn, Strings.T("nav.sidebarToggle"));
            btn.Click += (s, e) =>
            {
                try
                {
                    ToggleSidebar();
                }
                catch (Exception ex)
                {
                    Adapters.Logging.FileLogger.Error("ToggleSidebar", ex);
                }
            };
            return btn;
        }

        private void ToggleSidebar()
        {
            _sidebarOpen = !_sidebarOpen;
            if (_navHost != null)
            {
                _navHost.Visibility = _sidebarOpen ? Visibility.Visible : Visibility.Collapsed;
            }

            // glyph stays in the title row; state is just sidebar show/hide

            UpdatePanelWidths();
        }

        private void ShowPage(string page)
        {
            if (_navBusy)
            {
                return;
            }

            _navBusy = true;
            try
            {
                ShowPageCore(page);
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("ShowPage " + page, ex);
                if (_statusText != null)
                {
                    _statusText.Text = Strings.T("error.blank.prevented");
                }
            }
            finally
            {
                _navBusy = false;
            }
        }

        private void ShowPageCore(string page)
        {
            if (_quadrantPage != null)
            {
                _quadrantPage.Visibility = page == "quadrants" ? Visibility.Visible : Visibility.Collapsed;
            }

            if (_todoPage != null)
            {
                _todoPage.Visibility = page == "todo" ? Visibility.Visible : Visibility.Collapsed;
                if (page == "todo")
                {
                    BindTodos();
                }
            }

            if (_settingsPage != null)
            {
                _settingsPage.Visibility = page == "settings" ? Visibility.Visible : Visibility.Collapsed;
            }

            if (_navQuadrants != null)
            {
                StyleNav(_navQuadrants, page == "quadrants");
            }

            if (_navTodo != null)
            {
                StyleNav(_navTodo, page == "todo");
            }

            ApplyNavBadgeStyle(page == "todo");

            if (_navSettings != null)
            {
                StyleNav(_navSettings, page == "settings");
            }
        }

        private static void StyleNav(Button button, bool selected)
        {
            if (button == null)
            {
                return;
            }

            try
            {
                StyleNavCore(button, selected);
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("StyleNav", ex);
            }
        }

        private static void StyleNavCore(Button button, bool selected)
        {

            if (selected)
            {
                button.Background = Theme.AccentBrush;
                button.Foreground = Brushes.White;
                button.FontWeight = FontWeights.Medium;
                button.Template = UiKit.FilledPillTemplate();
            }
            else
            {
                button.Background = Brushes.Transparent;
                button.Foreground = Theme.InkBrush;
                button.FontWeight = FontWeights.Regular;
                button.Template = UiKit.TintedGlassPillTemplate();
            }

            // Icon-only entries (the settings gear) paint with geometry, so the glyph
            // has to follow the colour the pill template would have applied to text.
            var glyph = button.Content as System.Windows.Shapes.Shape;
            if (glyph != null)
            {
                glyph.Fill = button.Foreground;
            }

            button.Opacity = 1.0;
        }

        private UIElement BuildNav()
        {
            // Floating sidebar (Liquid Glass navigation layer). A three-row grid keeps
            // the read-only note and the settings button parked at the bottom edge of
            // the panel, however tall the window is.
            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var stack = new StackPanel { Margin = new Thickness(2, 10, 2, 10) };
            stack.Children.Add(new TextBlock
            {
                Text = Strings.T("app.title"),
                FontSize = UiKit.TypeTitle2,
                FontWeight = FontWeights.Bold,
                Foreground = Theme.InkBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(6, 4, 6, 18)
            });
            _navQuadrants = NavButton(Strings.T("nav.quadrants"), true, () => ShowPage("quadrants"));
            stack.Children.Add(_navQuadrants);
            _navTodo = NavButton(BuildNavTodoContent(), false, () => ShowPage("todo"));
            AutomationProperties.SetName(_navTodo, Strings.T("nav.todo"));
            stack.Children.Add(_navTodo);
            Grid.SetRow(stack, 0);
            grid.Children.Add(stack);

            var footer = new StackPanel { Margin = new Thickness(2, 0, 2, 10) };
            footer.Children.Add(new TextBlock
            {
                Text = Strings.T("status.readonly"),
                FontSize = UiKit.TypeCaption,
                FontWeight = FontWeights.Medium,
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(6, 0, 6, 10)
            });
            _navSettings = NavIconButton(UiKit.GearGlyph(), Strings.T("nav.settings"), () => ShowPage("settings"));
            footer.Children.Add(_navSettings);
            Grid.SetRow(footer, 2);
            grid.Children.Add(footer);

            return UiKit.GlassPlate(grid, UiKit.RadiusSheet, new Thickness(14, 18, 14, 18));
        }

        /// <summary>Icon-only sidebar entry (settings gear). The label lives in the tooltip.</summary>
        private Button NavIconButton(UIElement glyph, string tooltip, Action onClick)
        {
            var button = new Button
            {
                Content = glyph,
                Height = 38,
                Margin = new Thickness(0),
                HorizontalContentAlignment = HorizontalAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                ToolTip = tooltip,
                Foreground = Theme.InkBrush,
                Template = UiKit.TintedGlassPillTemplate()
            };
            AutomationProperties.SetName(button, tooltip);
            AutomationProperties.SetHelpText(button, tooltip);
            button.Click += (s, e) =>
            {
                try
                {
                    onClick();
                }
                catch (Exception ex)
                {
                    Adapters.Logging.FileLogger.Error("NavClick", ex);
                }
            };
            return button;
        }

        private Button NavButton(object content, bool selected, Action onClick)
        {
            var button = new Button
            {
                Content = content,
                Height = 38,
                Margin = new Thickness(0, 0, 0, 8),
                HorizontalContentAlignment = HorizontalAlignment.Center,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                FontSize = UiKit.TypeSubhead,
                FontWeight = selected ? FontWeights.Medium : FontWeights.Regular,
                Foreground = selected ? Brushes.White : Theme.InkBrush,
                Template = selected
                    ? UiKit.ProminentGlassButtonTemplate()
                    : UiKit.TintedGlassPillTemplate()
            };
            button.Click += (s, e) =>
            {
                try
                {
                    onClick();
                }
                catch (Exception ex)
                {
                    Adapters.Logging.FileLogger.Error("NavClick", ex);
                }
            };
            return button;
        }

        private ComboBox BuildDaysPicker()
        {
            var box = UiKit.Select();
            WithName(box, Strings.T("settings.days"));
            box.Width = 96;
            box.Height = 32;
            foreach (var d in new[] { "7", "14", "30", "60", "90" })
            {
                box.Items.Add(d);
            }

            box.SelectedItem = "30";
            return box;
        }

        private Panel BuildCenter()
        {
            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // Large title + grouped toolbar capsules (iOS toolbar item groups)
            var header = new StackPanel { Margin = new Thickness(4, 4, 4, 18) };
            header.Children.Add(UiKit.TitleRow(_sidebarToggle ?? (_sidebarToggle = CreateSidebarToggle()), UiKit.Title(Strings.T("nav.quadrants"))));

            var daysLabel = UiKit.Caption(Strings.T("scope.days"));
            daysLabel.VerticalAlignment = VerticalAlignment.Center;
            daysLabel.Margin = new Thickness(2, 0, 8, 0);
            var daysGroup = UiKit.ToolGroup(
                daysLabel,
                UiKit.FieldChrome((_daysBox = BuildDaysPicker())));

            var actionGroup = UiKit.ToolGroup(
                (_scanButton = UiKit.Primary(Strings.T("action.scan"), OnScanClick)),
                (_cancelButton = UiKit.Secondary(Strings.T("action.cancel"), OnCancelClick)));
            _scanButton.Height = 36;
            _scanButton.MinWidth = 110;
            _cancelButton.Height = 36;
            _cancelButton.MinWidth = 96;
            _cancelButton.IsEnabled = false;

            header.Children.Add(UiKit.ToolbarGroups(daysGroup, actionGroup));
            grid.Children.Add(header);
            Grid.SetRow(header, 0);

            // iOS segmented control
            var segments = new List<UIElement>
            {
                UiKit.SegmentItem(Strings.T("filter.all"), true),
                UiKit.SegmentItem(Strings.T("quadrant.q1"), false),
                UiKit.SegmentItem(Strings.T("quadrant.q2"), false),
                UiKit.SegmentItem(Strings.T("quadrant.q3"), false),
                UiKit.SegmentItem(Strings.T("quadrant.q4"), false)
            };
            _quadrantFilter = new ListBox
            {
                Margin = new Thickness(2, 0, 2, 12),
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                ItemContainerStyle = UiKit.FlatListItemStyle("Content.Tag"),
                Height = 48
            };
            ScrollViewer.SetHorizontalScrollBarVisibility(_quadrantFilter, ScrollBarVisibility.Disabled);
            ScrollViewer.SetVerticalScrollBarVisibility(_quadrantFilter, ScrollBarVisibility.Disabled);
            WrapHorizontal(_quadrantFilter);
            foreach (var seg in segments)
            {
                _quadrantFilter.Items.Add(seg);
            }

            _quadrantFilter.SelectionChanged += (s, e) => ApplyFilterFromUi();
            var segmentHost = UiKit.GlassPlate(_quadrantFilter, UiKit.RadiusPanel, new Thickness(6), false);
            grid.Children.Add(segmentHost);
            Grid.SetRow(segmentHost, 1);

            var body = new Grid();
            _listPanel = BuildListPanel();
            _emptyPanel = BuildEmptyPanel();
            _errorPanel = BuildErrorPanel();
            body.Children.Add(_listPanel);
            body.Children.Add(_emptyPanel);
            body.Children.Add(_errorPanel);
            var bodyWrap = UiKit.InsetGroupCard(body);
            grid.Children.Add(bodyWrap);
            Grid.SetRow(bodyWrap, 2);

            // Default the quadrant filter to the first quadrant, and let that selection do the
            // page's first bind (BindListCore runs ShowEmpty for the pre-scan state).
            // This has to happen after BuildListPanel: setting SelectedIndex raises
            // SelectionChanged synchronously, the handler rebinds the mail list, and binding a
            // list that does not exist yet threw a NullReferenceException on every launch.
            _quadrantFilter.SelectedIndex = 1;

            var footer = new StackPanel { Margin = new Thickness(8, 14, 8, 0) };
            _progress = new ProgressBar { Height = 6, Visibility = Visibility.Collapsed };
            _statusText = new TextBlock
            {
                FontSize = UiKit.TypeFootnote,
                FontWeight = FontWeights.Medium,
                Foreground = Theme.MutedBrush,
                Margin = new Thickness(0, 8, 0, 0)
            };
            footer.Children.Add(_progress);
            footer.Children.Add(_statusText);
            grid.Children.Add(footer);
            Grid.SetRow(footer, 3);
            return grid;
        }

        private static void WrapHorizontal(ListBox list)
        {
            var factory = new FrameworkElementFactory(typeof(StackPanel));
            factory.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);
            list.ItemsPanel = new ItemsPanelTemplate(factory);
        }

        private Panel BuildListPanel()
        {
            _mailList = new ListBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                ItemContainerStyle = UiKit.FlatListItemStyle()
            };
            ScrollViewer.SetVerticalScrollBarVisibility(_mailList, ScrollBarVisibility.Auto);
            ScrollViewer.SetHorizontalScrollBarVisibility(_mailList, ScrollBarVisibility.Disabled);
            ScrollViewer.SetCanContentScroll(_mailList, true);
            _mailList.SelectionChanged += (s, e) =>
            {
                if (_filterUpdating) return;
                OnMailSelected();
            };

            // Explicit ScrollViewer so long mail lists can scroll inside the panel.
            var scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Content = _mailList
            };
            var host = new Grid();
            host.Children.Add(scroll);
            return host;
        }

        private Panel BuildEmptyPanel()
        {
            var panel = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            panel.Children.Add(new TextBlock
            {
                Text = Strings.T("empty.title"),
                FontSize = UiKit.TypeLargeTitle,
                FontWeight = FontWeights.SemiBold,
                Foreground = Theme.InkBrush,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 8)
            });
            panel.Children.Add(new TextBlock
            {
                Text = Strings.T("empty.body"),
                Foreground = Theme.MutedBrush,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 10)
            });
            var start = UiKit.Primary(Strings.T("empty.start"), OnScanClick);
            start.HorizontalAlignment = HorizontalAlignment.Center;
            panel.Children.Add(start);
            return panel;
        }

        private Panel BuildErrorPanel()
        {
            var panel = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                MaxWidth = 420
            };
            panel.Children.Add(new TextBlock
            {
                Text = Strings.T("error.title"),
                FontSize = UiKit.TypeLargeTitle,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Theme.Danger),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 8)
            });
            panel.Children.Add(new TextBlock
            {
                Name = "errorBody",
                Text = string.Empty,
                TextWrapping = TextWrapping.Wrap,
                Foreground = Theme.InkBrush,
                TextAlignment = TextAlignment.Center,
                Margin = new Thickness(0, 0, 0, 10)
            });
            var retry = UiKit.Primary(Strings.T("action.retry"), OnScanClick);
            retry.HorizontalAlignment = HorizontalAlignment.Center;
            panel.Children.Add(retry);
            return panel;
        }

        private UIElement BuildDetail()
        {
            var stack = new StackPanel();
            _detailTitle = new TextBlock
            {
                FontSize = UiKit.TypeTitle2,
                FontWeight = FontWeights.Bold,
                Foreground = Theme.InkBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 14),
                LineHeight = 34
            };
            _detailMeta = new TextBlock
            {
                FontSize = UiKit.TypeFootnote,
                FontWeight = FontWeights.Medium,
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 18,
                Margin = new Thickness(0, 0, 0, 10)
            };
            _reasonPanel = new StackPanel();
            _reclassifyButton = UiKit.Secondary(Strings.T("action.reclassify"), OnReclassifyClick);
            _reclassifyButton.HorizontalAlignment = HorizontalAlignment.Stretch;
            _reclassifyButton.Margin = new Thickness(0, 12, 0, 0);
            _reclassifyButton.IsEnabled = false;
            _addTodoButton = UiKit.Primary(Strings.T("action.addTodo"), OnAddTodoClick);
            _addTodoButton.HorizontalAlignment = HorizontalAlignment.Stretch;
            _addTodoButton.Margin = new Thickness(0, 8, 0, 0);
            _addTodoButton.IsEnabled = false;
            _openOutlookButton = UiKit.Secondary(Strings.T("action.openInOutlook"), OnOpenInOutlookClick);
            _openOutlookButton.HorizontalAlignment = HorizontalAlignment.Stretch;
            _openOutlookButton.Margin = new Thickness(0, 12, 0, 0);
            _openOutlookButton.IsEnabled = false;
            _aiAction = UiKit.Secondary(Strings.T("ai.action"), OnAiActionClick);
            _aiAction.HorizontalAlignment = HorizontalAlignment.Stretch;
            _aiAction.Margin = new Thickness(0, 8, 0, 0);
            _aiAction.IsEnabled = false;
            stack.Children.Add(_detailTitle);
            stack.Children.Add(_detailMeta);
            stack.Children.Add(UiKit.Subtitle(Strings.T("reason.title")));
            stack.Children.Add(_reasonPanel);
            _todoActions = BuildTodoActionsPanel();
            stack.Children.Add(_todoActions);
            stack.Children.Add(_openOutlookButton);
            stack.Children.Add(_reclassifyButton);
            stack.Children.Add(_addTodoButton);
            stack.Children.Add(_aiAction);

            var close = new Button
            {
                Content = "×",
                Width = 24,
                Height = 24,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Foreground = Theme.MutedBrush,
                FontSize = 16,
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, -4, -4, 0)
            };
            close.Click += (s, e) => HideDetail();
            var host = new Grid();
            host.Children.Add(UiKit.GlassPlate(stack, UiKit.RadiusCard, new Thickness(18, 16, 18, 16)));
            host.Children.Add(close);
            return host;
        }

        private Panel BuildTodoPage()
        {
            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            grid.Children.Add(UiKit.TitleRow(CreateSidebarToggle(), UiKit.Title(Strings.T("todo.title"))));
            Grid.SetRow(grid.Children[grid.Children.Count - 1], 0);

            var tools = new StackPanel { Margin = new Thickness(4, 0, 4, 14) };
            var inputRow = new DockPanel { LastChildFill = true, Margin = new Thickness(0, 0, 0, 8) };
            var addManual = UiKit.Secondary(Strings.T("action.addTodoManual"), OnAddManualTodoClick);
            DockPanel.SetDock(addManual, Dock.Right);
            inputRow.Children.Add(addManual);
            _todoInput = UiKit.Input(string.Empty);
            WithName(_todoInput, Strings.T("todo.add.hint"));
            _todoInput.Margin = new Thickness(0, 0, 8, 0);
            _todoInput.KeyDown += (s, e) =>
            {
                if (e.Key == Key.Enter)
                {
                    OnAddManualTodoClick(s, e);
                    e.Handled = true;
                }
            };
            inputRow.Children.Add(_todoInput);
            tools.Children.Add(inputRow);
            _todoFilter = UiKit.Select();
            WithName(_todoFilter, Strings.T("todo.filter.label"));
            _todoFilter.Width = 150;
            _todoFilter.Items.Add(Strings.T("todo.filter.open"));
            _todoFilter.Items.Add(Strings.T("todo.filter.done"));
            _todoFilter.Items.Add(Strings.T("todo.filter.all"));
            _todoFilter.SelectedIndex = 0;
            _todoFilter.SelectionChanged += (s, e) =>
            {
                var i = _todoFilter.SelectedIndex;
                _todoFilterMode = i == 1 ? "Done" : (i == 2 ? "All" : "Open");
                BindTodos();
            };
            tools.Children.Add(_todoFilter);
            tools.Children.Add(BuildTodoToolsRow());
            grid.Children.Add(tools);
            Grid.SetRow(tools, 1);

            var body = new Grid();
            _todoList = new ListBox
            {
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                // Each row is named after its follow-up title, so assistive tools say
                // something better than the container's type name.
                ItemContainerStyle = UiKit.FlatListItemStyle("Content.Tag.Title")
            };
            ScrollViewer.SetVerticalScrollBarVisibility(_todoList, ScrollBarVisibility.Auto);
            ScrollViewer.SetHorizontalScrollBarVisibility(_todoList, ScrollBarVisibility.Disabled);
            _todoEmpty = new TextBlock
            {
                Text = Strings.T("todo.empty.body"),
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(24, 0, 24, 0)
            };
            body.Children.Add(_todoList);
            body.Children.Add(_todoEmpty);
            grid.Children.Add(UiKit.Panel(body));
            Grid.SetRow(grid.Children[grid.Children.Count - 1], 2);
            return grid;
        }

        private Panel BuildSettingsPage()
        {
            var outer = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            var panel = new StackPanel { Margin = new Thickness(4, 8, 4, 16) };
            panel.Children.Add(UiKit.Subtitle(Strings.T("nav.settings")));

            panel.Children.Add(UiKit.Caption(Strings.T("settings.days")));
            _settingsDays = UiKit.Input((_settings != null ? _settings.ScanDays : 30).ToString());
            WithName(_settingsDays, Strings.T("settings.days"));
            panel.Children.Add(_settingsDays);
            panel.Children.Add(UiKit.Caption(Strings.T("settings.scanScope")));

            panel.Children.Add(UiKit.Caption(Strings.T("settings.language")));
            _settingsLanguage = UiKit.Select();
            WithName(_settingsLanguage, Strings.T("settings.language"));
            _settingsLanguage.Width = 180;
            _settingsLanguage.Items.Add("繁體中文");
            _settingsLanguage.Items.Add("简体中文");
            _settingsLanguage.Items.Add("English");
            var langIndex = 0;
            if (_settings != null)
            {
                if (_settings.ToUiLanguage() == UiLanguage.ZhCn)
                {
                    langIndex = 1;
                }
                else if (_settings.ToUiLanguage() == UiLanguage.EnUs)
                {
                    langIndex = 2;
                }
            }

            _settingsLanguage.SelectedIndex = langIndex;
            panel.Children.Add(_settingsLanguage);

            panel.Children.Add(UiKit.Caption(Strings.T("settings.autoRefresh")));
            _autoRefreshBox = UiKit.Select();
            WithName(_autoRefreshBox, Strings.T("settings.autoRefresh"));
            _autoRefreshBox.Width = 180;
            foreach (var minutes in AutoRefreshChoices)
            {
                _autoRefreshBox.Items.Add(minutes == 0
                    ? Strings.T("settings.autoRefresh.off")
                    : string.Format(Strings.T("settings.autoRefresh.minutes"), minutes));
            }

            _autoRefreshBox.SelectedIndex = AutoRefreshChoiceIndex(_settings != null
                ? _settings.AutoRefreshMinutes
                : AppSettings.DefaultAutoRefreshMinutes);
            panel.Children.Add(_autoRefreshBox);

            panel.Children.Add(UiKit.Caption(Strings.T("settings.urgentKeywords")));
            _settingsUrgent = UiKit.Input(JoinKeywords(_settings != null ? _settings.UrgentKeywords : RuleOptions.DefaultUrgentKeywords()));
            WithName(_settingsUrgent, Strings.T("settings.urgentKeywords"));
            panel.Children.Add(_settingsUrgent);

            panel.Children.Add(UiKit.Caption(Strings.T("settings.importantKeywords")));
            _settingsImportant = UiKit.Input(JoinKeywords(_settings != null ? _settings.ImportantKeywords : RuleOptions.DefaultImportantKeywords()));
            WithName(_settingsImportant, Strings.T("settings.importantKeywords"));
            panel.Children.Add(_settingsImportant);

            panel.Children.Add(UiKit.Caption(Strings.T("settings.vip")));
            _settingsVip = UiKit.Input(JoinKeywords(_settings != null ? _settings.VipAddresses : new List<string>()));
            WithName(_settingsVip, Strings.T("settings.vip"));
            panel.Children.Add(_settingsVip);

            var save = UiKit.Primary(Strings.T("settings.save"), OnSaveSettingsClick);
            save.HorizontalAlignment = HorizontalAlignment.Left;
            save.Margin = new Thickness(0, 12, 0, 24);
            panel.Children.Add(save);

            panel.Children.Add(UiKit.Subtitle(Strings.T("ai.title")));
            _aiEnabled = UiKit.AppleCheck(Strings.T("ai.enabled"));
            _aiEnabled.IsChecked = _settings != null && _settings.AiEnabled;
            panel.Children.Add(_aiEnabled);

            panel.Children.Add(UiKit.Caption(Strings.T("ai.providers")));
            var providerRow = new StackPanel { Orientation = Orientation.Horizontal };
            _aiProviders = UiKit.Select();
            WithName(_aiProviders, Strings.T("ai.providers"));
            _aiProviders.Width = 180;
            _aiProviders.SelectionChanged += (s, e) => OnAiProviderSelected();
            providerRow.Children.Add(_aiProviders);
            var addP = UiKit.Secondary(Strings.T("ai.addProvider"), OnAiProviderAdd);
            addP.Margin = new Thickness(8, 0, 0, 0);
            addP.Height = 32;
            addP.MinWidth = 64;
            providerRow.Children.Add(addP);
            var delP = UiKit.Secondary(Strings.T("ai.removeProvider"), OnAiProviderRemove);
            delP.Margin = new Thickness(8, 0, 0, 0);
            delP.Height = 32;
            delP.MinWidth = 64;
            providerRow.Children.Add(delP);
            panel.Children.Add(providerRow);

            panel.Children.Add(UiKit.Caption(Strings.T("ai.baseUrl")));
            _aiBaseUrl = UiKit.Input(_settings != null ? _settings.AiBaseUrl : string.Empty);
            WithName(_aiBaseUrl, Strings.T("ai.baseUrl"));
            panel.Children.Add(_aiBaseUrl);

            panel.Children.Add(UiKit.Caption(Strings.T("ai.apiKey")));
            _aiApiKey = UiKit.Input(_settings != null ? _settings.AiApiKey : string.Empty);
            WithName(_aiApiKey, Strings.T("ai.apiKey"));
            panel.Children.Add(_aiApiKey);

            panel.Children.Add(UiKit.Caption(Strings.T("ai.models")));
            var modelRow = new StackPanel { Orientation = Orientation.Horizontal };
            _aiModel = UiKit.Select();
            WithName(_aiModel, Strings.T("ai.models"));
            _aiModel.Width = 220;
            modelRow.Children.Add(_aiModel);
            var loadBtn = UiKit.Secondary(Strings.T("ai.loadModels"), OnAiLoadModels);
            loadBtn.Margin = new Thickness(8, 0, 0, 0);
            loadBtn.Height = 32;
            loadBtn.MinWidth = 88;
            modelRow.Children.Add(loadBtn);
            panel.Children.Add(modelRow);
            _aiModelHint = new TextBlock
            {
                Text = Strings.T("ai.models.empty"),
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.MutedBrush,
                Margin = new Thickness(2, 6, 0, 12)
            };
            panel.Children.Add(_aiModelHint);

            panel.Children.Add(new TextBlock
            {
                Text = Strings.T("ai.key.local"),
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 4, 0, 20)
            });

            panel.Children.Add(UiKit.Subtitle(Strings.T("privacy.title")));
            panel.Children.Add(new TextBlock
            {
                Text = Strings.T("privacy.minimize"),
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });
            panel.Children.Add(new TextBlock
            {
                Text = Strings.T("privacy.ai.off"),
                Foreground = Theme.MutedBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 12)
            });
            panel.Children.Add(UiKit.Caption(Strings.T("privacy.location")));
            panel.Children.Add(new TextBlock
            {
                Text = JsonPaths.RootDirectory,
                FontFamily = new FontFamily("Consolas"),
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.InkBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 12)
            });

            var privacyButtons = new StackPanel { Orientation = Orientation.Horizontal };
            privacyButtons.Children.Add(UiKit.Primary(Strings.T("privacy.export"), OnExportClick));
            var openFolder = UiKit.Secondary(Strings.T("privacy.openFolder"), OnOpenExportFolderClick);
            openFolder.Margin = new Thickness(8, 0, 8, 0);
            privacyButtons.Children.Add(openFolder);
            privacyButtons.Children.Add(UiKit.Destructive(Strings.T("privacy.clear"), OnClearClick));
            panel.Children.Add(privacyButtons);

            panel.Children.Add(BuildUpdateSection());

            outer.Content = UiKit.Panel(panel);
            var host = new Grid();
            host.Children.Add(outer);
            return host;
        }

        /// <summary>
        /// Gives a field its own readable name. A caption sitting next to an input is invisible
        /// to a screen reader, which would otherwise announce a bare "edit" or "combo box".
        /// </summary>
        private static TextBox WithName(TextBox field, string name)
        {
            AutomationProperties.SetName(field, name);
            return field;
        }

        private static ComboBox WithName(ComboBox field, string name)
        {
            AutomationProperties.SetName(field, name);
            return field;
        }

        private static string JoinKeywords(IList<string> values)
        {
            return values == null ? string.Empty : string.Join(", ", values);
        }

        private static List<string> SplitKeywords(string text)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(text))
            {
                return list;
            }

            foreach (var part in text.Split(new[] { ',', '，', ';', '；' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmed = part.Trim();
                if (trimmed.Length > 0)
                {
                    list.Add(trimmed);
                }
            }

            return list;
        }

        private void BindTodos()
        {
            try
            {
                BindTodosCore();
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("BindTodos", ex);
            }
        }

        private void BindTodosCore()
        {
            if (_todoList == null)
            {
                return;
            }

            _todoList.Items.Clear();
            var filtered = _todos.Items.Where(t =>
                t != null &&
                (
                    _todoFilterMode == "All"
                    || (_todoFilterMode == "Done" && t.Status == TodoStatus.Done)
                    || (_todoFilterMode == "Open" && t.Status == TodoStatus.Open)
                ));

            // Ordering lives in Core/Models/TodoOrdering so "what has been waiting longest"
            // is testable without Outlook or a window; Manual is the stored order.
            var items = TodoOrdering.Sort(filtered, _todoSortMode, DateTime.UtcNow);

            foreach (var item in items)
            {
                _todoList.Items.Add(BuildTodoRow(item));
            }

            var empty = items.Count == 0;
            _todoEmpty.Visibility = empty ? Visibility.Visible : Visibility.Collapsed;
            _todoList.Visibility = empty ? Visibility.Collapsed : Visibility.Visible;
            UpdateTodoBadge();
        }

        private UIElement BuildTodoRow(TodoItem item)
        {
            var now = DateTime.UtcNow;
            // The row the detail panel is showing keeps a soft accent wash.
            var showing = !string.IsNullOrEmpty(_detailTodoId) && item.Id == _detailTodoId;
            // A previewed row wins over the overdue wash, so the panel's row is always
            // the one that stands out.
            var background = showing
                ? Theme.RowSelectedBrush
                : (item.IsOverdue(now) ? Theme.OverdueRowBrush : Brushes.Transparent);
            var grid = new Grid
            {
                Tag = item,
                MinHeight = 56,
                Background = background,
                Cursor = Cursors.Hand
            };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var toggle = new CheckBox
            {
                IsChecked = item.Status == TodoStatus.Done,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(12, 0, 8, 0)
            };
            // The box itself carries no text, so without this a screen reader would read a
            // bare "check box" and the user could not tell which task it completes.
            AutomationProperties.SetName(toggle, string.Format(Strings.T("todo.toggleDone"), item.Title));
            AutomationProperties.SetHelpText(toggle, Strings.T("todo.toggleDone.help"));
            toggle.Checked += (s, e) =>
            {
                if (_suppressTodoUi) return;
                if (item.Status != TodoStatus.Done)
                {
                    _suppressTodoUi = true;
                    try { _todos.ToggleDone(item.Id); BindTodos(); }
                    finally { _suppressTodoUi = false; }
                }
            };
            toggle.Unchecked += (s, e) =>
            {
                if (_suppressTodoUi) return;
                if (item.Status == TodoStatus.Done)
                {
                    _suppressTodoUi = true;
                    try { _todos.ToggleDone(item.Id); BindTodos(); }
                    finally { _suppressTodoUi = false; }
                }
            };
            Grid.SetColumn(toggle, 0);
            grid.Children.Add(toggle);

            var text = new StackPanel { Margin = new Thickness(4, 10, 8, 10), VerticalAlignment = VerticalAlignment.Center };
            var title = new TextBlock
            {
                Text = item.Title,
                FontWeight = FontWeights.SemiBold,
                Foreground = item.Status == TodoStatus.Done ? Theme.MutedBrush : Theme.InkBrush,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            if (item.Status == TodoStatus.Done)
            {
                title.TextDecorations = TextDecorations.Strikethrough;
            }

            text.Children.Add(title);
            if (!string.IsNullOrEmpty(item.SourceEntryId))
            {
                text.Children.Add(new TextBlock
                {
                    Text = string.Format(Strings.T("todo.source"), item.QuadrantHint ?? "-"),
                    FontSize = UiKit.TypeCaption,
                    Foreground = Theme.MutedBrush
                });
            }

            // Waiting / overdue sits under the source line, in the colour that says how
            // urgent it is: amber once it has waited a week, red once the due date passed.
            var status = WaitingStatusText(item, now);
            if (!string.IsNullOrEmpty(status))
            {
                text.Children.Add(new TextBlock
                {
                    Text = status,
                    FontSize = UiKit.TypeCaption,
                    FontWeight = FontWeights.Medium,
                    Foreground = WaitingStatusBrush(item, now),
                    TextTrimming = TextTrimming.CharacterEllipsis
                });
            }

            Grid.SetColumn(text, 1);
            grid.Children.Add(text);

            var del = UiKit.Secondary(Strings.T("action.todoDelete"), (s, e) =>
            {
                _todos.Remove(item.Id);
                BindTodos();
            });
            del.Margin = new Thickness(0, 0, 12, 0);
            del.MinWidth = 72;
            Grid.SetColumn(del, 2);
            grid.Children.Add(del);

            // Clicking the row itself (not its check box or delete button) previews the
            // mail behind the follow-up in the right-hand panel.
            grid.MouseLeftButtonUp += (s, e) =>
            {
                if (IsFromInteractiveChild(e.OriginalSource as DependencyObject))
                {
                    return;
                }

                ShowTodoDetail(item);
            };
            return grid;
        }

        private void OnAddTodoClick(object sender, RoutedEventArgs e)
        {
            try
            {
                OnAddTodoClickCore(sender, e);
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("OnAddTodoClick", ex);
            }
        }

        private void OnAddTodoClickCore(object sender, RoutedEventArgs e)
        {
            if (_selected == null)
            {
                return;
            }

            _todos.AddFromMail(_selected.Mail, Strings.QuadrantName(_selected.Classification.Quadrant));
            _statusText.Text = Strings.T("todo.added");
            ShowPage("todo");
        }

        private void OnAddManualTodoClick(object sender, RoutedEventArgs e)
        {
            if (_todoInput == null)
            {
                return;
            }

            var title = _todoInput.Text;
            if (string.IsNullOrEmpty(title) || string.IsNullOrEmpty(title.Trim()))
            {
                return;
            }

            _todos.AddManual(title);
            _todoInput.Text = string.Empty;
            BindTodos();
        }

        private void ShowEmpty()
        {
            if (_listPanel != null)
            {
                _listPanel.Visibility = Visibility.Collapsed;
            }

            if (_errorPanel != null)
            {
                _errorPanel.Visibility = Visibility.Collapsed;
            }

            if (_emptyPanel != null)
            {
                _emptyPanel.Visibility = Visibility.Visible;
            }
        }

        private void ShowList()
        {
            if (_emptyPanel != null)
            {
                _emptyPanel.Visibility = Visibility.Collapsed;
            }

            if (_errorPanel != null)
            {
                _errorPanel.Visibility = Visibility.Collapsed;
            }

            if (_listPanel != null)
            {
                _listPanel.Visibility = Visibility.Visible;
            }
        }

        private void ShowError(string key)
        {
            if (_listPanel != null)
            {
                _listPanel.Visibility = Visibility.Collapsed;
            }

            if (_emptyPanel != null)
            {
                _emptyPanel.Visibility = Visibility.Collapsed;
            }

            if (_errorPanel != null)
            {
                _errorPanel.Visibility = Visibility.Visible;
                var stack = _errorPanel as StackPanel;
                if (stack != null)
                {
                    foreach (var child in stack.Children.OfType<TextBlock>())
                    {
                        if (child.Name == "errorBody")
                        {
                            child.Text = Strings.T(key);
                        }
                    }
                }
            }

            _statusText.Text = Strings.T("error.blank.prevented");
        }

        private void ApplyFilterFromUi()
        {
            // SelectionChanged <-> Items/SelectedIndex can recurse until StackOverflow.
            if (_filterUpdating || _quadrantFilter == null)
            {
                return;
            }

            _filterUpdating = true;
            try
            {
                var index = _quadrantFilter.SelectedIndex;
                if (index < 0)
                {
                    index = 0;
                }

                if (index <= 0)
                {
                    _filterAll = true;
                }
                else
                {
                    _filterAll = false;
                    _filter = (Quadrant)(index - 1);
                }

                // Refresh segment visuals without touching SelectedIndex.
                var items = _quadrantFilter.Items;
                for (var i = 0; i < items.Count; i++)
                {
                    items[i] = UiKit.SegmentItem(GetSegmentLabel(i), index == i);
                }

                BindList();
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("ApplyFilterFromUi", ex);
            }
            finally
            {
                _filterUpdating = false;
            }
        }

        private string GetSegmentLabel(int index)
        {
            switch (index)
            {
                case 1:
                    return Strings.T("quadrant.q1");
                case 2:
                    return Strings.T("quadrant.q2");
                case 3:
                    return Strings.T("quadrant.q3");
                case 4:
                    return Strings.T("quadrant.q4");
                default:
                    return Strings.T("filter.all");
            }
        }

        private void BindList()
        {
            if (_navBusy) { return; }
            try
            {
                BindListCore();
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("BindList", ex);
                ShowEmpty();
            }
        }

        private void BindListCore()
        {
            // The page's first bind comes from the filter's initial selection; if anything ever
            // rebinds before BuildListPanel has run, there is simply nothing to bind into yet.
            if (_mailList == null)
            {
                return;
            }

            var source = _items
                .Where(i => _filterAll || i.Classification.Quadrant == _filter)
                .OrderByDescending(i => i.Classification.UrgencyScore + i.Classification.ImportanceScore)
                .ThenByDescending(i => i.Mail.ReceivedOn)
                .ToList();

            _mailList.Items.Clear();

            if (source.Count > 0)
            {
                var stack = UiKit.ListRows();
                for (var i = 0; i < source.Count; i++)
                {
                    UiKit.AddSeparated(stack, BuildRow(source[i]), i == source.Count - 1);
                }

                _mailList.Items.Add(UiKit.InsetGroupCard(stack));
            }

            if (_items.Count == 0)
            {
                ShowEmpty();
            }
            else
            {
                ShowList();
            }
        }

        private FrameworkElement BuildRow(ScanResultItem item)
        {
            var leading = UiKit.StatusDot(item.Classification.Quadrant);

            var titles = new StackPanel();
            titles.Children.Add(new TextBlock
            {
                Text = item.Mail.Subject,
                FontSize = UiKit.TypeBody,
                FontWeight = FontWeights.Regular,
                Foreground = Theme.InkBrush,
                TextTrimming = TextTrimming.CharacterEllipsis
            });
            titles.Children.Add(new TextBlock
            {
                Text = item.Mail.FromName + " · " + item.Mail.ReceivedOn.ToString("MM/dd HH:mm"),
                FontSize = UiKit.TypeCaption,
                FontWeight = FontWeights.Regular,
                Foreground = Theme.MutedBrush,
                Margin = new Thickness(0, 4, 0, 0),
                TextTrimming = TextTrimming.CharacterEllipsis
            });

            var trailing = new TextBlock
            {
                Text = Strings.QuadrantName(item.Classification.Quadrant),
                FontSize = UiKit.TypeCaption,
                FontWeight = FontWeights.Medium,
                Foreground = Theme.SecondaryBrush
            };

            var row = (System.Windows.Controls.Grid)UiKit.ListRow(leading, titles, trailing);
            row.Tag = item;
            row.Cursor = Cursors.Hand;
            row.Background = Brushes.Transparent;
            row.MouseLeftButtonUp += (s, e) =>
            {
                _lastPicked = item;
                OnMailSelectedDetail(item);
            };
            return row;
        }

        private void OnMailSelectedDetail(ScanResultItem item)
        {
            try
            {
                OnMailSelectedDetailCore(item);
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("OnMailSelectedDetail", ex);
            }
        }

        private void OnMailSelectedDetailCore(ScanResultItem item)
        {
            // Selecting a plain mail turns the follow-up block off again; ShowTodoDetail
            // re-renders it right after this call.
            SetTodoActionsVisible(false);
            _selected = item;
            if (item != null)
            {
                _detailEntryId = null;
            }
            if (_detailHost != null)
            {
                _detailHost.Visibility = item == null ? Visibility.Collapsed : Visibility.Visible;
            }

            UpdatePanelWidths();

            if (_reclassifyButton == null)
            {
                return;
            }

            if (item == null)
            {
                _reclassifyButton.IsEnabled = false;
                _addTodoButton.IsEnabled = false;
                if (_openOutlookButton != null) _openOutlookButton.IsEnabled = false;
                if (_aiAction != null)
                {
                    _aiAction.IsEnabled = false;
                }

                _detailTitle.Text = string.Empty;
                _detailMeta.Text = string.Empty;
                _reasonPanel.Children.Clear();
                return;
            }

            _reclassifyButton.IsEnabled = true;
            _addTodoButton.IsEnabled = true;
            if (_openOutlookButton != null) _openOutlookButton.IsEnabled = true;
            if (_aiAction != null)
            {
                _aiAction.IsEnabled = true;
            }

            _detailTitle.Text = item.Mail.Subject;
            _detailMeta.Text = Strings.T("detail.from") + ": " + item.Mail.FromName
                + "\n" + Strings.T("detail.time") + ": " + item.Mail.ReceivedOn.ToString("yyyy-MM-dd HH:mm")
                + "\n" + Strings.T("detail.folder") + ": " + item.Mail.FolderPath;

            _reasonPanel.Children.Clear();
            foreach (var reason in item.Classification.Reasons)
            {
                var label = Strings.T(reason.LabelKey);
                var detail = string.IsNullOrEmpty(reason.Detail) ? string.Empty : " (" + reason.Detail + ")";
                var delta = reason.Delta == 0 ? string.Empty : (reason.Delta > 0 ? " +" + reason.Delta : " " + reason.Delta);
                _reasonPanel.Children.Add(new TextBlock
                {
                    Text = "• " + label + detail + delta,
                    FontSize = UiKit.TypeSubhead,
                    Foreground = Theme.InkBrush,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, 8)
                });
            }
        }

        private void OnMailSelected()
        {
            _detailTodoId = null;
            _detailEntryId = null;
            var row = _mailList.SelectedItem as Grid;
            var picked = row == null ? (_lastPicked) : row.Tag as ScanResultItem;
            if (picked == null)
            {
                picked = _lastPicked;
            }

            OnMailSelectedDetail(picked);
        }


        private async void OnScanClick(object sender, RoutedEventArgs e)
        {
            try
            {
                await RunScanAsync();
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("OnScanClick", ex);
                ShowError("error.scan.failed");
            }
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            if (_cancel != null)
            {
                _cancel.Cancel();
            }
        }

        private async Task RunScanAsync()
        {
            int days;
            if (!int.TryParse(Convert.ToString(_daysBox != null ? _daysBox.SelectedItem : null), out days))
            {
                days = _settings != null ? _settings.ScanDays : 30;
            }

            _cancel = new CancellationTokenSourceLike();
            _scanRunning = true;
            _scanButton.IsEnabled = false;
            _cancelButton.IsEnabled = true;
            _progress.Visibility = Visibility.Visible;
            _progress.IsIndeterminate = true;
            _statusText.Text = Strings.T("loading.scanning");
            ShowList();
            Adapters.Logging.FileLogger.Info("UI scan requested days=" + days);

            var scope = new MailScanScope
            {
                Days = days,
                FolderPaths = new List<string> { _settings != null ? _settings.FolderPath : "Inbox" } // reader walks all stores
            };

            SyncRulesFromSettings(_settings);
            var useCase = new ScanMailUseCase(_reader, new RuleEngine(_ruleOptions), _overrideStore);
            ScanOutcome outcome = null;
            Exception failure = null;

            // COM must run on STA — never Task.Run/MTA (that hard-terminates the process).
            await Task.Run(() =>
            {
                try
                {
                    outcome = Adapters.Com.ComSta.Run(() => useCase.Run(
                        scope,
                        (done, total) =>
                        {
                            Dispatcher.Invoke(() =>
                            {
                                if (total > 0)
                                {
                                    _progress.IsIndeterminate = false;
                                    _progress.Maximum = total;
                                    _progress.Value = Math.Min(done, total);
                                }
                            });
                        },
                        () => _cancel.IsCancelled));
                }
                catch (Exception ex)
                {
                    failure = ex;
                    Adapters.Logging.FileLogger.Error("Scan pipeline crashed", ex);
                }
            });

            _scanButton.IsEnabled = true;
            _cancelButton.IsEnabled = false;
            _progress.Visibility = Visibility.Collapsed;
            _scanRunning = false;

            // Record what the list now corresponds to. This runs before the branches below
            // so that a failed scan cannot leave the probe comparing against a stale
            // fingerprint and rescanning on every tick.
            await CaptureInboxSignatureAsync();

            if (failure != null)
            {
                ShowError("error.scan.failed");
                return;
            }

            if (outcome == null)
            {
                ShowError("error.scan.failed");
                return;
            }

            if (!string.IsNullOrEmpty(outcome.ErrorKey))
            {
                ShowError(outcome.ErrorKey);
                return;
            }

            _items.Clear();
            foreach (var item in outcome.Items)
            {
                _items.Add(item);
            }

            _overrides.Clear();
            foreach (var pair in _overrideStore.Load())
            {
                _overrides[pair.Key] = pair.Value;
            }

            foreach (var item in _items)
            {
                OverrideEntry manual;
                if (_overrides.TryGetValue(item.Mail.EntryId, out manual) && manual != null)
                {
                    item.Classification = new ClassificationResult(
                        item.Mail.EntryId,
                        manual.Quadrant,
                        item.Classification.UrgencyScore,
                        item.Classification.ImportanceScore,
                        new List<ScoreReason>
                        {
                            new ScoreReason("override.manual", "reason.override.manual", 0, "Manual")
                        },
                        true);
                }
            }

            ApplyFilterFromUi();
            if (_items.Count == 0)
            {
                ShowEmpty();
                _statusText.Text = Strings.T("empty.title");
            }
            else if (outcome.Cancelled)
            {
                _statusText.Text = Strings.T("status.cancelled");
            }
            else
            {
                _statusText.Text = string.Format(Strings.T("status.done"), _items.Count);
            }
        }


        private void OnOpenInOutlookClick(object sender, RoutedEventArgs e)
        {
            try
            {
                // A follow-up shown without its scanned mail still knows its entry id.
                var entryId = _selected != null && !string.IsNullOrEmpty(_selected.Mail.EntryId)
                    ? _selected.Mail.EntryId
                    : _detailEntryId;
                if (string.IsNullOrEmpty(entryId))
                {
                    return;
                }

                var ok = _reader.TryOpenInOutlook(entryId);
                if (!ok)
                {
                    if (_statusText != null)
                    {
                        _statusText.Text = Strings.T("error.outlook.not_detected");
                    }
                }
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("OnOpenInOutlookClick", ex);
                if (_statusText != null)
                {
                    _statusText.Text = Strings.T("error.blank.prevented");
                }
            }
        }


        private void EnsureAiProviders()
        {
            if (_settings == null)
            {
                _settings = AppSettings.CreateDefault();
            }

            if (_settings.AiProviders == null)
            {
                _settings.AiProviders = new List<AiProviderProfile>();
            }

            if (_settings.AiProviders.Count == 0)
            {
                _settings.AiProviders.Add(new AiProviderProfile { Name = "Provider 1" });
            }

            if (string.IsNullOrEmpty(_settings.SelectedAiProviderId))
            {
                _settings.SelectedAiProviderId = _settings.AiProviders[0].Id;
            }
        }

        private void RefreshAiProviderCombo()
        {
            if (_aiProviders == null || _settings == null)
            {
                return;
            }

            EnsureAiProviders();
            _aiUiSync = true;
            try
            {
                _aiProviders.Items.Clear();
                foreach (var pr in _settings.AiProviders)
                {
                    var label = string.IsNullOrEmpty(pr.Name) ? "Provider" : pr.Name;
                    if (!string.IsNullOrEmpty(pr.BaseUrl))
                    {
                        label = label + " — " + pr.BaseUrl.Trim().TrimEnd('/');
                    }

                    _aiProviders.Items.Add(label);
                }

                var selected = _settings.GetSelectedProvider();
                var idx = selected != null ? _settings.AiProviders.IndexOf(selected) : 0;
                _aiProviders.SelectedIndex = idx < 0 ? 0 : idx;
                PushProviderToFields();
            }
            finally
            {
                _aiUiSync = false;
            }
        }

        private void PushProviderToFields()
        {
            var pr = _settings != null ? _settings.GetSelectedProvider() : null;
            if (_aiBaseUrl != null)
            {
                _aiBaseUrl.Text = pr != null ? pr.BaseUrl : string.Empty;
            }

            if (_aiApiKey != null)
            {
                _aiApiKey.Text = pr != null ? pr.ApiKey : string.Empty;
            }

            FillModelCombo(pr);
        }

        private void FillModelCombo(AiProviderProfile pr)
        {
            if (_aiModel == null)
            {
                return;
            }

            _aiUiSync = true;
            try
            {
                _aiModel.Items.Clear();
                if (pr != null && pr.CachedModels != null)
                {
                    foreach (var m in pr.CachedModels)
                    {
                        if (!string.IsNullOrEmpty(m))
                        {
                            _aiModel.Items.Add(m);
                        }
                    }
                }

                if (pr != null && !string.IsNullOrEmpty(pr.Model))
                {
                    if (!_aiModel.Items.Contains(pr.Model))
                    {
                        _aiModel.Items.Add(pr.Model);
                    }

                    _aiModel.SelectedItem = pr.Model;
                }
                else if (_aiModel.Items.Count > 0)
                {
                    _aiModel.SelectedIndex = 0;
                }
            }
            finally
            {
                _aiUiSync = false;
            }
        }

        private void SyncSelectedProviderFromUi()
        {
            if (_aiUiSync || _settings == null)
            {
                return;
            }

            EnsureAiProviders();
            var pr = _settings.GetSelectedProvider();
            if (pr == null)
            {
                return;
            }

            pr.BaseUrl = _aiBaseUrl != null ? _aiBaseUrl.Text.Trim() : pr.BaseUrl;
            pr.ApiKey = _aiApiKey != null ? _aiApiKey.Text.Trim() : pr.ApiKey;
            if (_aiModel != null && _aiModel.SelectedItem != null)
            {
                pr.Model = Convert.ToString(_aiModel.SelectedItem);
            }

            if (pr.CachedModels == null)
            {
                pr.CachedModels = new List<string>();
            }

            if (_aiModel != null)
            {
                foreach (var item in _aiModel.Items)
                {
                    var s = Convert.ToString(item);
                    if (!string.IsNullOrEmpty(s) && !pr.CachedModels.Contains(s))
                    {
                        pr.CachedModels.Add(s);
                    }
                }
            }
        }

        private void OnAiProviderSelected()
        {
            if (_aiUiSync || _aiProviders == null || _settings == null)
            {
                return;
            }

            SyncSelectedProviderFromUi();
            var idx = _aiProviders.SelectedIndex;
            EnsureAiProviders();
            if (idx >= 0 && idx < _settings.AiProviders.Count)
            {
                _settings.SelectedAiProviderId = _settings.AiProviders[idx].Id;
                PushProviderToFields();
            }
        }

        private void OnAiProviderAdd(object sender, RoutedEventArgs e)
        {
            EnsureAiProviders();
            SyncSelectedProviderFromUi();
            var n = _settings.AiProviders.Count + 1;
            var pr = new AiProviderProfile { Name = "Provider " + n };
            _settings.AiProviders.Add(pr);
            _settings.SelectedAiProviderId = pr.Id;
            RefreshAiProviderCombo();
        }

        private void OnAiProviderRemove(object sender, RoutedEventArgs e)
        {
            EnsureAiProviders();
            if (_settings.AiProviders.Count <= 1)
            {
                return;
            }

            var pr = _settings.GetSelectedProvider();
            if (pr == null)
            {
                return;
            }

            _settings.AiProviders.Remove(pr);
            _settings.SelectedAiProviderId = _settings.AiProviders[0].Id;
            RefreshAiProviderCombo();
        }

        private void OnAiLoadModels(object sender, RoutedEventArgs e)
        {
            try
            {
                SyncSelectedProviderFromUi();
                EnsureAiProviders();
                var pr = _settings.GetSelectedProvider();
                if (pr == null || string.IsNullOrEmpty(pr.BaseUrl) || string.IsNullOrEmpty(pr.ApiKey))
                {
                    if (_aiModelHint != null)
                    {
                        _aiModelHint.Text = Strings.T("ai.models.empty");
                    }

                    return;
                }

                var cap = new OutlookAiHelper.Adapters.Ai.HttpAiCapability(
                    new OutlookAiHelper.Adapters.Ai.AiProviderConfig
                    {
                        BaseUrl = pr.BaseUrl,
                        ApiKey = pr.ApiKey,
                        Model = pr.Model,
                        TimeoutMs = 15000
                    });
                var models = cap.ListModels();
                if (models == null || models.Count == 0)
                {
                    if (_aiModelHint != null)
                    {
                        _aiModelHint.Text = Strings.T("ai.models.failed");
                    }

                    return;
                }

                pr.CachedModels = models;
                FillModelCombo(pr);
                if (_aiModelHint != null)
                {
                    _aiModelHint.Text = string.Format(Strings.T("ai.models.loaded"), models.Count);
                }
            }
            catch (Exception ex)
            {
                Adapters.Logging.FileLogger.Error("OnAiLoadModels", ex);
                if (_aiModelHint != null)
                {
                    _aiModelHint.Text = Strings.T("ai.models.failed");
                }
            }
        }

        private void OnReclassifyClick(object sender, RoutedEventArgs e)
        {
            if (_selected == null)
            {
                return;
            }

            var dialog = new Window
            {
                Title = Strings.T("pick.quadrant"),
                Width = 320,
                Height = 300,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize,
                Background = Brushes.Transparent
            };

            var stack = new StackPanel { Margin = new Thickness(4) };
            foreach (var q in new[]
            {
                Quadrant.Q1UrgentImportant,
                Quadrant.Q2UrgentNotImportant,
                Quadrant.Q3ImportantNotUrgent,
                Quadrant.Q4Neither
            })
            {
                var target = q;
                var button = UiKit.Primary(Strings.QuadrantName(target), (s, args) =>
                {
                    ApplyOverride(_selected, target);
                    dialog.DialogResult = true;
                });
                button.Margin = new Thickness(0, 0, 0, 8);
                button.Height = 42;
                stack.Children.Add(button);
            }

            dialog.Content = UiKit.GlassPlate(stack, UiKit.RadiusSheet, new Thickness(28, 22, 28, 24));
            dialog.ShowDialog();
        }

        private void ApplyOverride(ScanResultItem item, Quadrant quadrant)
        {
            var entry = OverrideEntry.Manual(item.Mail.EntryId, quadrant);
            _overrides[item.Mail.EntryId] = entry;
            _overrideStore.Save(_overrides);
            item.Classification = new ClassificationResult(
                item.Mail.EntryId,
                quadrant,
                item.Classification.UrgencyScore,
                item.Classification.ImportanceScore,
                new List<ScoreReason>
                {
                    new ScoreReason("override.manual", "reason.override.manual", 0, "Manual")
                },
                true);
            BindList();
            OnMailSelected();
        }

        private void OnSaveSettingsClick(object sender, RoutedEventArgs e)
        {
            int days = 30;
            if (_settingsDays != null)
            {
                int.TryParse(_settingsDays.Text.Trim(), out days);
            }

            if (days <= 0)
            {
                days = 30;
            }

            if (_settings == null)
            {
                _settings = AppSettings.CreateDefault();
            }

            var previousLanguage = _settings.Language;
            _settings.ScanDays = days;
            _settings.UrgentKeywords = SplitKeywords(_settingsUrgent != null ? _settingsUrgent.Text : null);
            _settings.ImportantKeywords = SplitKeywords(_settingsImportant != null ? _settingsImportant.Text : null);
            _settings.VipAddresses = SplitKeywords(_settingsVip != null ? _settingsVip.Text : null);
            _settings.AiEnabled = _aiEnabled != null && _aiEnabled.IsChecked == true;
            SyncSelectedProviderFromUi();
            var selectedP = _settings.GetSelectedProvider();
            _settings.AiBaseUrl = selectedP != null ? selectedP.BaseUrl : string.Empty;
            _settings.AiModel = selectedP != null ? selectedP.Model : string.Empty;
            _settings.AiApiKey = selectedP != null ? selectedP.ApiKey : string.Empty;
            _settings.CheckForUpdatesOnStartup = _updateAutoCheck != null && _updateAutoCheck.IsChecked == true;
            _settings.AutoRefreshMinutes = SelectedAutoRefreshMinutes();
            var langIndex = _settingsLanguage != null ? _settingsLanguage.SelectedIndex : 0;
            _settings.Language = langIndex == 1 ? "zh-CN" : (langIndex == 2 ? "en-US" : "zh-TW");

            if (_settings.UrgentKeywords.Count == 0)
            {
                _settings.UrgentKeywords = RuleOptions.DefaultUrgentKeywords();
            }

            if (_settings.ImportantKeywords.Count == 0)
            {
                _settings.ImportantKeywords = RuleOptions.DefaultImportantKeywords();
            }

            _settingsStore.Save(_settings);
            SyncRulesFromSettings(_settings);
            ConfigureAutoRefresh();
            Strings.Language = _settings.ToUiLanguage();
            ApplySettingsToToolbar();
            RefreshAiProviderCombo();
            _statusText.Text = Strings.T("settings.saved");

            if (!string.Equals(previousLanguage, _settings.Language, StringComparison.OrdinalIgnoreCase))
            {
                ShowRestartDialog();
            }
        }

        private void ShowRestartDialog()
        {
            var dialog = new Window
            {
                Title = Strings.T("restart.title"),
                Width = 420,
                Height = 240,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize,
                Background = Brushes.Transparent
            };

            var stack = new StackPanel { Margin = new Thickness(24) };
            stack.Children.Add(new TextBlock
            {
                Text = Strings.T("restart.body"),
                FontSize = UiKit.TypeBody,
                Foreground = Theme.InkBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 20)
            });

            var row = new StackPanel { Orientation = Orientation.Horizontal };
            var now = UiKit.Primary(Strings.T("restart.now"), (s, e) =>
            {
                dialog.DialogResult = true;
                RestartApp();
            });
            var later = UiKit.Secondary(Strings.T("restart.later"), (s, e) =>
            {
                dialog.DialogResult = true;
                var exe = Process.GetCurrentProcess().MainModule.FileName;
                if (!string.IsNullOrEmpty(exe))
                {
                    try
                    {
                        Process.Start(exe);
                    }
                    catch (Exception)
                    {
                    }
                }
            });
            var next = UiKit.Secondary(Strings.T("restart.next"), (s, e) =>
            {
                dialog.DialogResult = true;
            });
            now.Margin = new Thickness(0, 0, 8, 0);
            later.Margin = new Thickness(0, 0, 8, 0);
            row.Children.Add(now);
            row.Children.Add(later);
            row.Children.Add(next);
            stack.Children.Add(row);
            dialog.Content = UiKit.GlassPlate(stack, UiKit.RadiusSheet, new Thickness(28, 22, 28, 24));
            dialog.ShowDialog();
        }

        private void OnAiActionClick(object sender, RoutedEventArgs e)
        {
            if (_selected == null)
            {
                return;
            }

            var ai = BuildAi();
            if (ai == null || !ai.IsEnabled)
            {
                _statusText.Text = Strings.T("ai.disabled");
                return;
            }

            var reasons = _selected.Classification.Reasons
                .Select(r => Strings.T(r.LabelKey))
                .ToList();
            var request = new AiRequest
            {
                Subject = _selected.Mail.Subject,
                FromName = _selected.Mail.FromName,
                Quadrant = Strings.QuadrantName(_selected.Classification.Quadrant),
                Reasons = reasons,
                Instruction = "Explain priority in 2 short sentences and suggest one next action."
            };

            if (!ConfirmAiWhitelist(request))
            {
                return;
            }

            try
            {
                var result = ai.Suggest(request);
                if (result != null && result.Success)
                {
                    MessageBox.Show(result.Text, Strings.T("ai.action"), MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    _statusText.Text = Strings.T(result != null && result.ErrorKey != null ? result.ErrorKey : "ai.failed");
                }
            }
            catch (Exception)
            {
                _statusText.Text = Strings.T("ai.failed");
            }
        }

        private IAiCapability BuildAi()
        {
            if (_settings == null || !_settings.AiEnabled)
            {
                return new OutlookAiHelper.Adapters.Ai.NullAiCapability();
            }

            var sel = _settings.GetSelectedProvider();
            return new OutlookAiHelper.Adapters.Ai.HttpAiCapability(new OutlookAiHelper.Adapters.Ai.AiProviderConfig
            {
                BaseUrl = sel != null ? sel.BaseUrl : _settings.AiBaseUrl,
                Model = sel != null ? sel.Model : _settings.AiModel,
                ApiKey = sel != null ? sel.ApiKey : _settings.AiApiKey,
                TimeoutMs = 15000
            });
        }

        private bool ConfirmAiWhitelist(AiRequest request)
        {
            var dialog = new Window
            {
                Title = Strings.T("ai.consent.title"),
                Width = 460,
                Height = 360,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize,
                Background = Brushes.Transparent
            };

            var stack = new StackPanel { Margin = new Thickness(24) };
            stack.Children.Add(new TextBlock
            {
                Text = Strings.T("ai.consent.body"),
                TextWrapping = TextWrapping.Wrap,
                Foreground = Theme.InkBrush,
                Margin = new Thickness(0, 0, 0, 12)
            });
            stack.Children.Add(new TextBlock
            {
                Text = "Subject: " + request.Subject
                    + "\nFrom: " + request.FromName
                    + "\nQuadrant: " + request.Quadrant
                    + "\nReasons: " + string.Join(", ", request.Reasons),
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new FontFamily("Consolas"),
                FontSize = UiKit.TypeCaption,
                Foreground = Theme.InkBrush,
                Margin = new Thickness(0, 0, 0, 10)
            });

            var approved = false;
            var row = new StackPanel { Orientation = Orientation.Horizontal };
            var send = UiKit.Primary(Strings.T("ai.consent.send"), (s, e) =>
            {
                approved = true;
                dialog.DialogResult = true;
            });
            var cancel = UiKit.Secondary(Strings.T("ai.consent.cancel"), (s, e) =>
            {
                dialog.DialogResult = true;
            });
            send.Margin = new Thickness(0, 0, 8, 0);
            row.Children.Add(send);
            row.Children.Add(cancel);
            stack.Children.Add(row);
            dialog.Content = UiKit.GlassPlate(stack, UiKit.RadiusSheet, new Thickness(28, 22, 28, 24));
            dialog.ShowDialog();
            return approved;
        }

        private void RestartApp()
        {
            try
            {
                var exe = Process.GetCurrentProcess().MainModule.FileName;
                if (!string.IsNullOrEmpty(exe))
                {
                    Process.Start(exe);
                }
            }
            catch (Exception)
            {
            }

            System.Windows.Application.Current.Shutdown();
        }

        private void OnExportClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var exportRoot = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "OutlookAiHelper");
                var path = _privacy.ExportAll(exportRoot);
                MessageBox.Show(Strings.T("privacy.exported") + path, Strings.T("app.title"), MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Strings.T("error.title") + "\n" + ex.Message, Strings.T("app.title"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnOpenExportFolderClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var exportRoot = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "OutlookAiHelper");
                Directory.CreateDirectory(exportRoot);
                Process.Start("explorer.exe", exportRoot);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Strings.T("error.title") + "\n" + ex.Message, Strings.T("app.title"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnClearClick(object sender, RoutedEventArgs e)
        {
            var confirm = MessageBox.Show(
                Strings.T("privacy.clear.confirm"),
                Strings.T("privacy.clear"),
                MessageBoxButton.OKCancel,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.OK)
            {
                return;
            }

            try
            {
                _privacy.ClearLocalData(true, true, true);
                _todos.ClearAll();
                _overrides.Clear();
                _items.Clear();
                _settings = AppSettings.CreateDefault();
                _settingsStore.Save(_settings);
                SyncRulesFromSettings(_settings);
                ApplySettingsToToolbar();
            RefreshAiProviderCombo();
                BindTodos();
                BindList();
                _statusText.Text = Strings.T("privacy.cleared");
                ShowPage("quadrants");
                ShowEmpty();
            }
            catch (Exception ex)
            {
                MessageBox.Show(Strings.T("error.title") + "\n" + ex.Message, Strings.T("app.title"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private sealed class CancellationTokenSourceLike
        {
            private volatile bool _cancelled;

            public bool IsCancelled
            {
                get { return _cancelled; }
            }

            public void Cancel()
            {
                _cancelled = true;
            }
        }
    }
}

