using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Shapes;
using Separator = System.Windows.Controls.Separator;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// Liquid Glass materials (Apple HIG) painted with layered brushes.
    /// Window chrome stays native Windows; every in-app control uses these tokens.
    /// </summary>
    public static class UiKit
    {
        // Concentric curvature (controls nest into container corners)
        public const double RadiusField = 6;
        public const double RadiusControl = 6;
        public const double RadiusCard = 10;
        public const double RadiusPanel = 10;
        public const double RadiusSection = 10;
        public const double RadiusPill = 999;
        public const double RadiusSheet = 10;

        // Spacing
        public const double Space1 = 8;
        public const double Space2 = 12;
        public const double Space3 = 16;
        public const double Space4 = 24;
        public const double Space5 = 32;

        // Type — refined compact SF scale (elegant, not oversized)
        public const double TypeLargeTitle = 22;
        public const double TypeTitle2 = 17;
        public const double TypeHeadline = 13;
        public const double TypeBody = 12;
        public const double TypeSubhead = 13;
        public const double TypeFootnote = 12;
        public const double TypeCaption = 11;
        public const double TypeButton = 13;
        public const double TypeNav = 13;

        public static CornerRadius R(double r)
        {
            return new CornerRadius(r);
        }

        #region Liquid Glass material

        /// <summary>Glass fill — quiet white with soft cool tint.</summary>
        public static Brush LiquidGlassFill(double opacity = 1.0)
        {
            // macOS translucent material (light) — shows canvas through
            var b = new LinearGradientBrush
            {
                StartPoint = new Point(0.5, 0),
                EndPoint = new Point(0.5, 1)
            };
            b.GradientStops.Add(new GradientStop(Color.FromArgb(A(0xD8, opacity), 0xFF, 0xFF, 0xFF), 0.0));
            b.GradientStops.Add(new GradientStop(Color.FromArgb(A(0xC0, opacity), 0xF8, 0xF8, 0xFA), 0.55));
            b.GradientStops.Add(new GradientStop(Color.FromArgb(A(0xCC, opacity), 0xF2, 0xF2, 0xF5), 1.0));
            return b;
        }

        /// <summary>Clearer glass for fields.</summary>
        public static Brush LiquidGlassClear()
        {
            return new SolidColorBrush(Color.FromArgb(0xFF, 0xE8, 0xE6, 0xE1));
        }

        /// <summary>Edge stroke — soft white rim.</summary>
        public static Brush GlassEdge()
        {
            return new SolidColorBrush(Color.FromArgb(0xFF, 0xE8, 0xE6, 0xE1));
        }

        /// <summary>Specular sheen (very subtle).</summary>
        public static Brush GlassSheen()
        {
            var b = new LinearGradientBrush
            {
                StartPoint = new Point(0, 0),
                EndPoint = new Point(0.2, 1)
            };
            b.GradientStops.Add(new GradientStop(Color.FromArgb(0x22, 0xFF, 0xFF, 0xFF), 0.0));
            b.GradientStops.Add(new GradientStop(Color.FromArgb(0x00, 0xFF, 0xFF, 0xFF), 0.55));
            return b;
        }

        /// <summary>Top specular hairline.</summary>
        public static Brush GlassTopSpec()
        {
            return new LinearGradientBrush(
                Color.FromArgb(0xC0, 0xFF, 0xFF, 0xFF),
                Color.FromArgb(0x00, 0xFF, 0xFF, 0xFF),
                90);
        }

        /// <summary>Prominent glass — muted sage accent (Morandi).</summary>
        public static Brush ProminentGlass()
        {
            var b = new LinearGradientBrush
            {
                StartPoint = new Point(0.5, 0),
                EndPoint = new Point(0.5, 1)
            };
            b.GradientStops.Add(new GradientStop(Color.FromRgb(0x2B, 0x88, 0xD8), 0.0));
            b.GradientStops.Add(new GradientStop(Color.FromRgb(0x00, 0x78, 0xD4), 1.0));
            return b;
        }

        public static Brush CardBrush()
        {
            return new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF));
        }

        public static Brush FillSecondary()
        {
            return new SolidColorBrush(Color.FromRgb(0xE0, 0xDD, 0xD7));
        }

        public static Brush SeparatorBrush()
        {
            return new SolidColorBrush(Color.FromArgb(0x33, 0xB8, 0xB0, 0xA4));
        }

        public static Brush FillPrimary()
        {
            return ProminentGlass();
        }

        private static byte A(byte a, double opacity)
        {
            var v = (int)(a * opacity);
            if (v < 0)
            {
                v = 0;
            }

            if (v > 255)
            {
                v = 255;
            }

            return (byte)v;
        }

        #endregion

        /// <summary>Soft float shadow.</summary>
        public static DropShadowEffect GlassFloatShadow()
        {
            return new DropShadowEffect
            {
                BlurRadius = 18,
                ShadowDepth = 4,
                Opacity = 0.10,
                Direction = 270,
                Color = Color.FromRgb(0x3A, 0x3A, 0x3C)
            };
        }

        public static DropShadowEffect ControlShadow()
        {
            return new DropShadowEffect
            {
                BlurRadius = 8,
                ShadowDepth = 1,
                Opacity = 0.08,
                Direction = 270,
                Color = Color.FromRgb(0x3A, 0x3A, 0x3C)
            };
        }

        /// <summary>
        /// Floating Liquid Glass plate: fill + sheen + edge + top spec + inner bottom light.
        /// Used for nav, toolbars, and sheets — the "controls layer" above content.
        /// </summary>
        public static Border GlassPlate(UIElement child, double radius, Thickness padding, bool prominentShadow = true)
        {
            var body = new Border
            {
                Background = LiquidGlassFill(),
                BorderBrush = GlassEdge(),
                BorderThickness = new Thickness(1.15),
                CornerRadius = R(radius),
                Child = new Border
                {
                    Padding = padding,
                    Background = Brushes.Transparent,
                    Child = child
                }
            };

            var sheen = new Border
            {
                Background = GlassSheen(),
                CornerRadius = R(radius),
                IsHitTestVisible = false
            };

            var topSpec = new Border
            {
                Height = 1.5,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(radius * 0.4, 1, radius * 0.4, 0),
                Background = GlassTopSpec(),
                CornerRadius = new CornerRadius(1),
                IsHitTestVisible = false
            };

            var bottomLift = new Border
            {
                Margin = new Thickness(1.5),
                BorderThickness = new Thickness(0, 0, 0, 1.2),
                BorderBrush = new LinearGradientBrush(
                    Color.FromArgb(0x00, 0xFF, 0xFF, 0xFF),
                    Color.FromArgb(0x55, 0xFF, 0xFF, 0xFF),
                    90),
                CornerRadius = new CornerRadius(0, 0, radius, radius),
                IsHitTestVisible = false
            };

            var stack = new Grid();
            stack.Children.Add(body);
            stack.Children.Add(sheen);
            stack.Children.Add(topSpec);
            stack.Children.Add(bottomLift);

            return new Border
            {
                                Child = stack
            };
        }

        /// <summary>Content card (list section) — white, large radius, not glass.</summary>
        public static Border ContentCard(UIElement child, double radius, Thickness padding)
        {
            return new Border
            {
                Background = CardBrush(),
                BorderBrush = SeparatorBrush(),
                BorderThickness = new Thickness(1),
                CornerRadius = R(radius),
                Padding = padding,
                Child = child
            };
        }

        public static Border Island(UIElement child, double radius, Brush material, Thickness? padding = null)
        {
            // Prefer glass plates for chrome; content uses ContentCard via callers.
            return GlassPlate(child, radius, padding ?? new Thickness(Space3));
        }

        public static Border Panel(UIElement child)
        {
            return GlassPlate(child, RadiusSheet, new Thickness(Space4, Space3, Space4, Space3));
        }

        public static Border Card(UIElement child)
        {
            return ContentCard(child, RadiusCard, new Thickness(Space3, Space2, Space3, Space2));
        }

        public static Border Toolbar(UIElement child)
        {
            return GlassPlate(child, RadiusControl, new Thickness(Space2, Space1, Space2, Space1));
        }

        /// <summary>
        /// iOS toolbar item group — one glass capsule shared by related controls
        /// (e.g. Scan + Cancel together; days picker as its own group).
        /// </summary>
        public static Border ToolGroup(params UIElement[] items)
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center
            };
            panel.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            for (var i = 0; i < items.Length; i++)
            {
                if (items[i] != null)
                {
                    var el = items[i];
                    if (el is FrameworkElement)
                    {
                        ((FrameworkElement)el).Margin = i == 0
                            ? new Thickness(0)
                            : new Thickness(6, 0, 0, 0);
                    }

                    panel.Children.Add(el);
                }
            }

            return GlassPlate(panel, RadiusControl, new Thickness(8, 6, 8, 6), false);
        }

        /// <summary>Toolbar bar containing grouped glass capsules with spacing between groups.</summary>
        public static Border ToolbarGroups(params Border[] groups)
        {
            var bar = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center
            };
            foreach (var g in groups)
            {
                if (g == null)
                {
                    continue;
                }

                if (bar.Children.Count > 0)
                {
                    g.Margin = new Thickness(Space2, 0, 0, 0);
                }

                bar.Children.Add(g);
            }

            return new Border
            {
                Child = bar,
                Background = Brushes.Transparent
            };
        }

        /// <summary>Context menu / action sheet chrome — Liquid Glass list.</summary>
        public static ContextMenu GlassMenu()
        {
            var menu = new ContextMenu
            {
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(6),
                FontSize = TypeSubhead,
                Foreground = Theme.InkBrush
            };
            menu.Template = MenuChromeTemplate();
            menu.ItemContainerStyle = MenuItemStyle();
            return menu;
        }

        private static Style MenuItemStyle()
        {
            var style = new Style(typeof(MenuItem));
            style.Setters.Add(new Setter(Control.FontSizeProperty, TypeSubhead));
            style.Setters.Add(new Setter(Control.ForegroundProperty, Theme.InkBrush));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(16, 10, 16, 10)));
            style.Setters.Add(new Setter(Control.BackgroundProperty, Brushes.Transparent));
            style.Setters.Add(new Setter(Control.TemplateProperty, MenuItemTemplate()));
            return style;
        }

        private static ControlTemplate MenuItemTemplate()
        {
            var template = new ControlTemplate(typeof(MenuItem));
            var root = new FrameworkElementFactory(typeof(Border));
            root.Name = "Bd";
            root.SetValue(Border.CornerRadiusProperty, R(12));
            root.SetValue(Border.BackgroundProperty, Brushes.Transparent);
            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            root.AppendChild(content);
            template.VisualTree = root;

            var hover = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
            hover.Setters.Add(new Setter(Border.BackgroundProperty, FillSecondary(), "Bd"));
            template.Triggers.Add(hover);
            return template;
        }

        private static ControlTemplate MenuChromeTemplate()
        {
            var template = new ControlTemplate(typeof(ContextMenu));
            var plate = new FrameworkElementFactory(typeof(Border));
            plate.SetValue(Border.CornerRadiusProperty, R(RadiusControl));
            plate.SetValue(Border.BackgroundProperty, LiquidGlassFill(0.98));
            plate.SetValue(Border.BorderBrushProperty, (Brush)new SolidColorBrush(Color.FromRgb(0xD0, 0xCB, 0xC3)));
            plate.SetValue(Border.BorderThicknessProperty, new Thickness(1.1));
            plate.SetValue(Border.PaddingProperty, new Thickness(2, 6, 2, 6));

            var sheen = new FrameworkElementFactory(typeof(Border));
            sheen.SetValue(Border.CornerRadiusProperty, R(RadiusControl));
            sheen.SetValue(Border.BackgroundProperty, GlassSheen());
            sheen.SetValue(UIElement.IsHitTestVisibleProperty, false);

            var presenter = new FrameworkElementFactory(typeof(ItemsPresenter));
            var stack = new FrameworkElementFactory(typeof(Grid));
            stack.AppendChild(plate);
            stack.AppendChild(sheen);
            stack.AppendChild(presenter);
            template.VisualTree = stack;
            return template;
        }

        /// <summary>iOS inset grouped list: section header + rounded white card.</summary>
        public static FrameworkElement InsetGroup(string sectionHeader, UIElement rows)
        {
            var stack = new StackPanel();
            if (!string.IsNullOrEmpty(sectionHeader))
            {
                stack.Children.Add(new TextBlock
                {
                    Text = sectionHeader,
                    FontSize = TypeFootnote,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(Theme.SecondaryLabel),
                    Margin = new Thickness(14, 0, 0, 8)
                });
            }

            stack.Children.Add(ContentCard(rows, RadiusSection, new Thickness(0)));
            return stack;
        }

        /// <summary>Row container that draws iOS hairline separators between children.</summary>
        public static StackPanel ListRows()
        {
            return new StackPanel();
        }

        /// <summary>iOS inset grouped list card with hairline separators.</summary>
        public static Border InsetGroupCard(UIElement rows)
        {
            return new Border
            {
                Background = CardBrush(),
                BorderBrush = new SolidColorBrush(Color.FromArgb(0x1A, 0xB8, 0xB0, 0xA4)),
                BorderThickness = new Thickness(1),
                CornerRadius = R(RadiusPanel),
                ClipToBounds = true,
                Child = rows
            };
        }

        /// <summary>One list row — 56pt iOS content row.</summary>
        public static FrameworkElement ListRow(UIElement leading, UIElement title, UIElement trailing)
        {
            var grid = new Grid { MinHeight = 56 };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            if (leading != null)
            {
                ((FrameworkElement)leading).Margin = new Thickness(16, 0, 10, 0);
                ((FrameworkElement)leading).VerticalAlignment = VerticalAlignment.Center;
                Grid.SetColumn((UIElement)leading, 0);
                grid.Children.Add(leading);
            }

            if (title != null)
            {
                ((FrameworkElement)title).Margin = leading == null
                    ? new Thickness(16, 12, 10, 12)
                    : new Thickness(0, 12, 10, 12);
                ((FrameworkElement)title).VerticalAlignment = VerticalAlignment.Center;
                Grid.SetColumn((UIElement)title, 1);
                grid.Children.Add(title);
            }

            if (trailing != null)
            {
                ((FrameworkElement)trailing).Margin = new Thickness(8, 0, 16, 0);
                ((FrameworkElement)trailing).VerticalAlignment = VerticalAlignment.Center;
                Grid.SetColumn((UIElement)trailing, 2);
                grid.Children.Add(trailing);
            }

            return grid;
        }

        public static void AddSeparated(StackPanel list, FrameworkElement row, bool isLast)
        {
            list.Children.Add(row);
            if (!isLast)
            {
                list.Children.Add(new Separator
                {
                    Height = 1,
                    Background = new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF)),
                    Margin = new Thickness(16, 0, 0, 0)
                });
            }
        }

        /// <summary>Half / full glass sheet dialog chrome (inset from window edges).</summary>
        public static Window GlassSheet(string title, UIElement content, double width, double height)
        {
            var sheet = new Window
            {
                Title = title,
                Width = width,
                Height = height,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                Background = Brushes.Transparent,
                WindowStyle = WindowStyle.SingleBorderWindow,
                ShowInTaskbar = false
            };

            // Outer margin creates the "inset half sheet" feel against the parent
            var outer = new Border
            {
                Margin = new Thickness(20, 24, 20, 24),
                Child = GlassPlate(content, RadiusSheet, new Thickness(28, 22, 28, 24))
            };

            sheet.Content = outer;
            return sheet;
        }

        /// <summary>Action-sheet style vertical stack of glass buttons.</summary>
        public static StackSheetBuilder ActionSheet()
        {
            return new StackSheetBuilder();
        }

        public sealed class StackSheetBuilder
        {
            private readonly StackPanel _stack = new StackPanel();

            public StackSheetBuilder Text(string text, bool headline)
            {
                _stack.Children.Add(new TextBlock
                {
                    Text = text,
                    FontSize = headline ? TypeHeadline : TypeBody,
                    FontWeight = headline ? FontWeights.SemiBold : FontWeights.Regular,
                    Foreground = Theme.InkBrush,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 0, 0, headline ? 12 : 16)
                });
                return this;
            }

            public StackSheetBuilder Button(Button button)
            {
                if (button != null)
                {
                    button.Margin = new Thickness(0, 0, 0, 10);
                    button.HorizontalAlignment = HorizontalAlignment.Stretch;
                    _stack.Children.Add(button);
                }

                return this;
            }

            public UIElement Build()
            {
                return _stack;
            }
        }

        #region Controls — glass button styles

        /// <summary>UIButton.Configuration.glassProminent()</summary>
        private static DataTemplate NoWrapContentTemplate()
        {
            // Must use a real Binding — SetValue("{Binding}") stores a literal string.
            var factory = new FrameworkElementFactory(typeof(TextBlock));
            factory.SetValue(TextBlock.TextWrappingProperty, TextWrapping.NoWrap);
            factory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.None);
            factory.SetValue(TextBlock.FontSizeProperty, TypeButton);
            factory.SetValue(TextBlock.FontWeightProperty, FontWeights.Medium);
            factory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            factory.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            factory.SetBinding(TextBlock.TextProperty, new Binding());
            return new DataTemplate { VisualTree = factory };
        }

        public static Button Primary(string text, RoutedEventHandler onClick)
        {
            var button = new Button
            {
                Content = text,
                Height = 28,
                MinWidth = 80,
                Padding = new Thickness(16, 0, 16, 0),
                Background = Brushes.Transparent,
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                FontSize = TypeButton,
                FontWeight = FontWeights.Medium,
                ContentTemplate = NoWrapContentTemplate()
            };
            button.Template = ProminentGlassButtonTemplate();
            button.Click += onClick;
            return button;
        }

        /// <summary>UIButton.Configuration.glass()</summary>
        public static Button Destructive(string text, RoutedEventHandler onClick)
        {
            // macOS destructive action: neutral push with red label/border (not a huge colored pill)
            var button = new Button
            {
                Content = text,
                Height = 28,
                MinWidth = 80,
                Padding = new Thickness(16, 0, 16, 0),
                Background = new SolidColorBrush(Color.FromRgb(0xF5, 0xF3, 0xEF)),
                Foreground = new SolidColorBrush(Color.FromRgb(0xC0, 0x39, 0x2B)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xC0, 0x39, 0x2B)),
                BorderThickness = new Thickness(1),
                Cursor = Cursors.Hand,
                FontSize = TypeButton,
                FontWeight = FontWeights.Medium,
                Template = DestructivePillTemplate()
            };
            button.Click += onClick;
            return button;
        }

        private static ControlTemplate DestructivePillTemplate()
        {
            var template = new ControlTemplate(typeof(Button));
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty, R(6));
            border.SetValue(Border.BackgroundProperty, new SolidColorBrush(Color.FromRgb(0xF5, 0xF3, 0xEF)));
            border.SetValue(Border.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(0xC0, 0x39, 0x2B)));
            border.SetValue(Border.BorderThicknessProperty, new Thickness(1));
            border.SetValue(Border.PaddingProperty, new Thickness(14, 0, 14, 0));
            var presenter = new FrameworkElementFactory(typeof(ContentPresenter));
            presenter.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            presenter.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            presenter.SetValue(System.Windows.Documents.TextElement.ForegroundProperty, new SolidColorBrush(Color.FromRgb(0xC0, 0x39, 0x2B)));
            presenter.SetValue(System.Windows.Documents.TextElement.FontSizeProperty, TypeButton);
            presenter.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            presenter.SetValue(ContentPresenter.RecognizesAccessKeyProperty, false);
            presenter.SetValue(System.Windows.Documents.TextElement.FontSizeProperty, TypeButton);
            presenter.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            presenter.SetValue(ContentPresenter.RecognizesAccessKeyProperty, false);
            border.AppendChild(presenter);
            template.VisualTree = border;
            return template;
        }

        public static Button Secondary(string text, RoutedEventHandler onClick)
        {
            var button = new Button
            {
                Content = text,
                Height = 28,
                MinWidth = 80,
                Padding = new Thickness(16, 0, 16, 0),
                Background = Brushes.Transparent,
                Foreground = Theme.InkBrush,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                FontSize = TypeButton,
                FontWeight = FontWeights.Regular,
                ContentTemplate = NoWrapContentTemplate()
            };
            button.Template = GlassButtonTemplate(false);
            button.Click += onClick;
            return button;
        }

        public static ControlTemplate ProminentGlassButtonTemplate()
        {
            return GlassButtonTemplate(true);
        }

        public static ControlTemplate FilledPillTemplate()
        {
            return GlassButtonTemplate(true);
        }

        public static ControlTemplate TintedGlassPillTemplate()
        {
            return GlassButtonTemplate(false);
        }

        public static ControlTemplate SimplePillTemplate(Brush background, Brush foreground)
        {
            return GlassButtonTemplate(background == ProminentGlass() || (background is SolidColorBrush && ((SolidColorBrush)background).Color == Theme.Accent));
        }

        public static ControlTemplate GlassButtonTemplatePublic(bool prominent)
        {
            return GlassButtonTemplate(prominent);
        }

        private static ControlTemplate GlassButtonTemplate(bool prominent)
        {
            var template = new ControlTemplate(typeof(Button));
            var root = new FrameworkElementFactory(typeof(Grid));

            var plate = new FrameworkElementFactory(typeof(Border));
            plate.SetValue(Border.CornerRadiusProperty, R(RadiusControl));
            plate.SetValue(Border.BackgroundProperty, prominent ? ProminentGlass() : (Brush)new SolidColorBrush(Color.FromRgb(0xE2, 0xDF, 0xD9)));
            plate.SetValue(Border.BorderBrushProperty, prominent ? (Brush)new SolidColorBrush(Color.FromArgb(0x55, 0xFF, 0xFF, 0xFF)) : (Brush)new SolidColorBrush(Color.FromRgb(0xC5, 0xC0, 0xB8)));
            plate.SetValue(Border.BorderThicknessProperty, new Thickness(1));
            
            var sheen = new FrameworkElementFactory(typeof(Border));
            sheen.SetValue(Border.CornerRadiusProperty, R(RadiusControl));
            sheen.SetValue(Border.BackgroundProperty, GlassSheen());
            sheen.SetValue(UIElement.IsHitTestVisibleProperty, false);

            var top = new FrameworkElementFactory(typeof(Border));
            top.SetValue(Border.HeightProperty, 1.25);
            top.SetValue(Border.VerticalAlignmentProperty, VerticalAlignment.Top);
            top.SetValue(Border.MarginProperty, new Thickness(16, 1, 16, 0));
            top.SetValue(Border.BackgroundProperty, GlassTopSpec());
            top.SetValue(UIElement.IsHitTestVisibleProperty, false);

            var presenter = new FrameworkElementFactory(typeof(ContentPresenter));
            presenter.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            presenter.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            presenter.SetValue(System.Windows.Documents.TextElement.ForegroundProperty, prominent ? Brushes.White : Theme.InkBrush);

            root.AppendChild(plate);
            root.AppendChild(sheen);
            root.AppendChild(top);
            root.AppendChild(presenter);
            template.VisualTree = root;
            return template;
        }

        /// <summary>UITextField — glass field.</summary>
        public static CheckBox AppleCheck(string content)
        {
            return new CheckBox
            {
                Content = content,
                FontSize = TypeBody,
                FontWeight = FontWeights.Regular,
                Foreground = Theme.InkBrush,
                VerticalContentAlignment = VerticalAlignment.Center,
                MinHeight = 24,
                Margin = new Thickness(0, 2, 0, 8)
            };
        }

        public static TextBox Input(string text)
        {
            var box = new TextBox
            {
                Text = text ?? string.Empty,
                Height = 36,
                Background = Brushes.Transparent,
                Foreground = Theme.InkBrush,
                BorderThickness = new Thickness(0),
                FontSize = TypeBody,
                FontWeight = FontWeights.Regular,
                CaretBrush = Theme.AccentBrush,
                VerticalContentAlignment = VerticalAlignment.Center
            };
            box.Template = GlassInputTemplate();
            return box;
        }

        public static ControlTemplate GlassInputTemplate()
        {
            var template = new ControlTemplate(typeof(TextBox));
            var root = new FrameworkElementFactory(typeof(Grid));

            var plate = new FrameworkElementFactory(typeof(Border));
            plate.SetValue(Border.CornerRadiusProperty, R(RadiusField));
            plate.SetValue(Border.BackgroundProperty, (Brush)new SolidColorBrush(Color.FromRgb(0xF2, 0xF0, 0xEB)));
            plate.SetValue(Border.BorderBrushProperty, (Brush)new SolidColorBrush(Color.FromRgb(0xD0, 0xCB, 0xC3)));
            plate.SetValue(Border.BorderThicknessProperty, new Thickness(1));
            plate.SetValue(Border.PaddingProperty, new Thickness(Space2, 0, Space2, 0));

            var sheen = new FrameworkElementFactory(typeof(Border));
            sheen.SetValue(Border.CornerRadiusProperty, R(RadiusField));
            sheen.SetValue(Border.BackgroundProperty, GlassSheen());
            sheen.SetValue(UIElement.IsHitTestVisibleProperty, false);

            var scroll = new FrameworkElementFactory(typeof(ScrollViewer));
            scroll.Name = "PART_ContentHost";
            scroll.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);

            root.AppendChild(plate);
            root.AppendChild(sheen);
            root.AppendChild(scroll);
            template.VisualTree = root;
            return template;
        }

        public static ControlTemplate CapsuleInputTemplate()
        {
            return GlassInputTemplate();
        }

        /// <summary>
        /// iOS picker field: glass chrome around the STOCK ComboBox control.
        /// No custom Binding/TemplateBinding on Popup — those hard-crashed startup.
        /// </summary>
        public static ComboBox Select()
        {
            var box = new ComboBox
            {
                Height = 32,
                MinWidth = 88,
                FontSize = TypeSubhead,
                FontWeight = FontWeights.Regular,
                Foreground = Theme.InkBrush,
                Padding = new Thickness(10, 0, 22, 0),
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0)
            };

            try
            {
                box.Template = (ControlTemplate)XamlReader.Parse(GlassComboXaml());
                box.ItemContainerStyle = MorandiComboItemStyle();
            }
            catch (Exception)
            {
                // keep stock template — dropdown still works
            }

            return box;
        }

        private static Style MorandiComboItemStyle()
        {
            // macOS menu item: focus/selection = inverted filled background
            var style = new Style(typeof(ComboBoxItem));
            style.Setters.Add(new Setter(Control.FontSizeProperty, TypeSubhead));
            style.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.Regular));
            style.Setters.Add(new Setter(Control.ForegroundProperty, Theme.InkBrush));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(12, 7, 12, 7)));
            style.Setters.Add(new Setter(Control.MarginProperty, new Thickness(4, 1, 4, 1)));
            style.Setters.Add(new Setter(Control.BackgroundProperty, Brushes.Transparent));
            style.Setters.Add(new Setter(Control.TemplateProperty, MacMenuItemTemplate()));
            return style;
        }

        private static ControlTemplate MacMenuItemTemplate()
        {
            var template = new ControlTemplate(typeof(ComboBoxItem));
            var root = new FrameworkElementFactory(typeof(Border));
            root.Name = "Bd";
            root.SetValue(Border.CornerRadiusProperty, R(6));
            root.SetValue(Border.BackgroundProperty, Brushes.Transparent);

            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            content.SetValue(ContentPresenter.MarginProperty, new Thickness(2, 0, 2, 0));
            root.AppendChild(content);
            template.VisualTree = root;

            // hover — soft fill
            var hover = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
            hover.Setters.Add(new Setter(Border.BackgroundProperty, new SolidColorBrush(Color.FromArgb(0x22, 0x00, 0x00, 0x00)), "Bd"));
            hover.Setters.Add(new Setter(Control.ForegroundProperty, Theme.InkBrush));
            template.Triggers.Add(hover);

            // keyboard / mouse focus — macOS highlight (inverted)
            var focus = new Trigger { Property = ComboBoxItem.IsHighlightedProperty, Value = true };
            focus.Setters.Add(new Setter(Border.BackgroundProperty, new SolidColorBrush(Color.FromArgb(0xE6, 0x00, 0x7A, 0xFF)), "Bd"));
            focus.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
            template.Triggers.Add(focus);

            // committed selection — inverted, slightly deeper
            var selected = new Trigger { Property = ComboBoxItem.IsSelectedProperty, Value = true };
            selected.Setters.Add(new Setter(Border.BackgroundProperty, new SolidColorBrush(Color.FromArgb(0xF2, 0x00, 0x7A, 0xFF)), "Bd"));
            selected.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
            selected.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.Medium));
            template.Triggers.Add(selected);

            return template;
        }

        private static string GlassComboXaml()
        {
            // macOS vibrancy: translucent popover, focus ring, inverted menu items
            return @"<ControlTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
                     xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
                     TargetType='{x:Type ComboBox}'>
  <Grid>
    <ToggleButton x:Name='ToggleButton' Focusable='False' ClickMode='Press'
                  IsChecked='{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}, Mode=TwoWay}'>
      <ToggleButton.Template>
        <ControlTemplate TargetType='{x:Type ToggleButton}'>
          <Border x:Name='Chrome' CornerRadius='7' BorderThickness='1'
                  BorderBrush='#D0CBC3' Background='#F2F0EB'>
            <Border.Effect>
              <DropShadowEffect BlurRadius='8' ShadowDepth='1' Opacity='0.10' Color='#3A4A5A'/>
            </Border.Effect>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property='IsMouseOver' Value='True'>
              <Setter TargetName='Chrome' Property='Background' Value='#CCFFFFFF'/>
            </Trigger>
            <Trigger Property='IsChecked' Value='True'>
              <Setter TargetName='Chrome' Property='Background' Value='#E6FFFFFF'/>
              <Setter TargetName='Chrome' Property='BorderBrush' Value='#990078D4'/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </ToggleButton.Template>
    </ToggleButton>
    <ContentPresenter Margin='10,0,22,0' VerticalAlignment='Center' HorizontalAlignment='Left'
                      IsHitTestVisible='False'
                      Content='{TemplateBinding SelectionBoxItem}'
                      ContentTemplate='{TemplateBinding SelectionBoxItemTemplate}'
                      ContentTemplateSelector='{TemplateBinding ItemTemplateSelector}'/>
    <Path HorizontalAlignment='Right' VerticalAlignment='Center' Margin='0,0,10,0'
          Width='8' Height='4.5' Fill='#6E6E73' Stretch='Uniform'
          Data='M 0,0 L 4,4 L 8,0 L 7,0 L 4,3 L 1,0 Z' IsHitTestVisible='False'/>
    <Popup x:Name='PART_Popup' Placement='Bottom' HorizontalOffset='0' VerticalOffset='4'
           AllowsTransparency='True' PopupAnimation='Fade'
           IsOpen='{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}}'>
      <Border CornerRadius='8' Padding='4' MinWidth='104'
              Background='#FBFAF8' BorderBrush='#D0CBC3' BorderThickness='1'>
        <Border.Effect>
          <DropShadowEffect BlurRadius='28' ShadowDepth='8' Opacity='0.22' Color='#1A2230'/>
        </Border.Effect>
        <ScrollViewer MaxHeight='280' VerticalScrollBarVisibility='Auto' HorizontalScrollBarVisibility='Disabled'>
          <ItemsPresenter/>
        </ScrollViewer>
      </Border>
    </Popup>
  </Grid>
</ControlTemplate>";
        }

        public static Border FieldChrome(UIElement inner)
        {
            return new Border
            {
                CornerRadius = R(10),
                Background = new SolidColorBrush(Color.FromRgb(0xF2, 0xF0, 0xEB)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xD0, 0xCB, 0xC3)),
                BorderThickness = new Thickness(1),
                MinHeight = 32,
                Padding = new Thickness(2, 0, 2, 0),
                Child = inner
            };
        }

        private static Style ComboItemStyle()
        {
            // Stock item visuals — no custom templates, no bindings (startup safety).
            var style = new Style(typeof(ComboBoxItem));
            style.Setters.Add(new Setter(Control.FontSizeProperty, TypeSubhead));
            style.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.Regular));
            style.Setters.Add(new Setter(Control.ForegroundProperty, Theme.InkBrush));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(12, 8, 12, 8)));
            style.Setters.Add(new Setter(Control.BackgroundProperty, Brushes.Transparent));
            return style;
        }

        /// <summary>UISegmentedControl item — compact pill only (no ListBox chrome).</summary>
        public static Border SegmentItem(string text, bool selected)
        {
            var label = new TextBlock
            {
                Text = text,
                FontSize = TypeSubhead,
                FontWeight = selected ? FontWeights.Medium : FontWeights.Regular,
                Foreground = selected ? Brushes.White : new SolidColorBrush(Theme.SecondaryLabel),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            return new Border
            {
                Margin = new Thickness(2, 0, 2, 0),
                Padding = new Thickness(11, 5, 11, 5),
                CornerRadius = R(6),
                Background = selected ? new SolidColorBrush(Color.FromRgb(0x00, 0x7A, 0xFF)) : Brushes.Transparent,
                                Cursor = Cursors.Hand,
                Child = label
            };
        }

        /// <summary>Horizontal segmented strip — no ListBox selection frame.</summary>
        public static StackPanel SegmentStrip(params UIElement[] items)
        {
            var strip = new StackPanel
            {
                Orientation = Orientation.Horizontal
            };
            foreach (var item in items)
            {
                strip.Children.Add(item);
            }

            return strip;
        }

        /// <summary>ListBoxItem container that draws content only (kills oversized selection frame).</summary>
        public static Style FlatListItemStyle()
        {
            var style = new Style(typeof(ListBoxItem));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(0)));
            style.Setters.Add(new Setter(Control.MarginProperty, new Thickness(0)));
            style.Setters.Add(new Setter(Control.BackgroundProperty, Brushes.Transparent));
            style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0)));
            style.Setters.Add(new Setter(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Stretch));
            style.Setters.Add(new Setter(Control.TemplateProperty, ContentOnlyItemTemplate()));
            return style;
        }

        private static ControlTemplate ContentOnlyItemTemplate()
        {
            // Group/list container: content only. Selection invert lives on combo menu items.
            var template = new ControlTemplate(typeof(ListBoxItem));
            var root = new FrameworkElementFactory(typeof(ContentPresenter));
            root.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
            root.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            template.VisualTree = root;
            return template;
        }

        #endregion

        /// <summary>Soft float shadow.</summary>
        /// <summary>macOS sidebar.left SF-symbol-like glyph.</summary>
        public static UIElement SidebarGlyph(double size = 14)
        {
            var canvas = new Canvas
            {
                Width = size,
                Height = size,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            var frame = new System.Windows.Shapes.Rectangle
            {
                Width = size - 1,
                Height = size - 3,
                Stroke = Theme.MutedBrush,
                StrokeThickness = 1.2,
                RadiusX = 1.5,
                RadiusY = 1.5
            };
            Canvas.SetLeft(frame, 0.5);
            Canvas.SetTop(frame, 1.5);
            canvas.Children.Add(frame);
            var divider = new System.Windows.Shapes.Line
            {
                X1 = size * 0.38,
                Y1 = 1.5,
                X2 = size * 0.38,
                Y2 = size - 1.5,
                Stroke = Theme.MutedBrush,
                StrokeThickness = 1.2
            };
            canvas.Children.Add(divider);
            return canvas;
        }

        /// <summary>
        /// macOS sidebar gear glyph (settings). Drawn as geometry so it stays crisp
        /// at small sizes and needs no icon font. The caller owns the colour: the
        /// sidebar restyles Fill when the settings page is selected.
        /// </summary>
        public static System.Windows.Shapes.Path GearGlyph(double size = 16, Brush brush = null)
        {
            var centre = size / 2.0;
            var bandOuter = size * 0.30;
            var bandInner = size * 0.115;
            var toothOuter = size * 0.475;
            var toothInner = size * 0.22;
            var toothWidth = size * 0.16;

            var group = new GeometryGroup { FillRule = FillRule.Nonzero };
            // Hub: an excluded ring, so the teeth can merge into it without notches.
            group.Children.Add(new CombinedGeometry(
                GeometryCombineMode.Exclude,
                new EllipseGeometry(new Point(centre, centre), bandOuter, bandOuter),
                new EllipseGeometry(new Point(centre, centre), bandInner, bandInner)));

            for (var i = 0; i < 8; i++)
            {
                var tooth = new RectangleGeometry(
                    new Rect(centre - (toothWidth / 2.0), centre - toothOuter, toothWidth, toothOuter - toothInner),
                    size * 0.04,
                    size * 0.04);
                tooth.Transform = new RotateTransform(i * 45.0, centre, centre);
                group.Children.Add(tooth);
            }

            return new System.Windows.Shapes.Path
            {
                Data = group,
                Fill = brush ?? Theme.InkBrush,
                Width = size,
                Height = size,
                Stretch = Stretch.None,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
        }

        public static TextBlock Caption(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = TypeCaption,
                FontWeight = FontWeights.Regular,
                Foreground = new SolidColorBrush(Theme.TertiaryLabel),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(2, 0, 8, 0)
            };
        }

        public static UIElement TitleRow(UIElement leading, TextBlock title)
        {
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(28) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            if (leading != null)
            {
                ((FrameworkElement)leading).VerticalAlignment = VerticalAlignment.Center;
                ((FrameworkElement)leading).HorizontalAlignment = HorizontalAlignment.Left;
                Grid.SetColumn((UIElement)leading, 0);
                grid.Children.Add(leading);
            }
            if (title != null)
            {
                title.VerticalAlignment = VerticalAlignment.Center;
                title.Margin = new Thickness(4, 0, 0, 0);
                Grid.SetColumn(title, 1);
                grid.Children.Add(title);
            }
            return grid;
        }

        public static TextBlock Title(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = TypeLargeTitle,
                FontWeight = FontWeights.SemiBold,
                Foreground = Theme.InkBrush,
                Margin = new Thickness(2, 0, 0, 10),
                LineHeight = 28
            };
        }

        public static TextBlock Subtitle(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = TypeTitle2,
                FontWeight = FontWeights.SemiBold,
                Foreground = Theme.InkBrush,
                Margin = new Thickness(2, 0, 0, 10)
            };
        }

        public static TextBlock Body(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontSize = TypeBody,
                FontWeight = FontWeights.Regular,
                Foreground = Theme.InkBrush,
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 20
            };
        }

        public static Ellipse StatusDot(Core.Classification.Quadrant q)
        {
            return new Ellipse
            {
                Width = 10,
                Height = 10,
                Fill = Theme.QuadrantBrush(q),
                VerticalAlignment = VerticalAlignment.Center
            };
        }

        public static Rectangle AccentBar(Core.Classification.Quadrant q)
        {
            return new Rectangle
            {
                Fill = Theme.QuadrantBrush(q),
                Width = 4,
                RadiusX = 2,
                RadiusY = 2
            };
        }
    }
}
