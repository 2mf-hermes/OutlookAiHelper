using System.Windows;
using System.Windows.Media;

namespace OutlookAiHelper.UI
{
    /// <summary>
    /// Morandi + Apple-inspired tokens. Quiet, professional, no AI-cliché neon.
    /// Window chrome is native Windows; everything inside uses this system.
    /// </summary>
    public static class Theme
    {
        // Morandi neutrals
        public static readonly Color Canvas = Color.FromRgb(0xEC, 0xEA, 0xE6);
        public static readonly Color CanvasTop = Color.FromRgb(0xF3, 0xF1, 0xED);
        public static readonly Color Surface = Color.FromRgb(0xF9, 0xF7, 0xF4);
        public static readonly Color Card = Color.FromRgb(0xFE, 0xFD, 0xFB);
        public static readonly Color Label = Color.FromRgb(0x3F, 0x3B, 0x38);
        public static readonly Color SecondaryLabel = Color.FromRgb(0x6F, 0x69, 0x63);
        public static readonly Color TertiaryLabel = Color.FromRgb(0x9A, 0x93, 0x8B);
        public static readonly Color Separator = Color.FromRgb(0xD9, 0xD4, 0xCD);
        public static readonly Color Fill = Color.FromRgb(0xE6, 0xE2, 0xDB);

        // Accents — Outlook-aligned blue + Morandi neutrals
        public static readonly Color Accent = Color.FromRgb(0x00, 0x78, 0xD4); // Microsoft Outlook / Fluent
        public static readonly Color AccentDeep = Color.FromRgb(0x03, 0x64, 0xB8);
        public static readonly Color AccentSoft = Color.FromRgb(0x8F, 0xA6, 0xB5); // blue-gray
        public static readonly Color Rose = Color.FromRgb(0xC4, 0xA5, 0x9B);
        public static readonly Color Gold = Color.FromRgb(0xC2, 0xA8, 0x78);

        public static readonly Color Q1 = Color.FromRgb(0xC4, 0x8A, 0x82);
        public static readonly Color Q2 = Color.FromRgb(0xC2, 0xA8, 0x78);
        public static readonly Color Q3 = Color.FromRgb(0x00, 0x78, 0xD4); // important = Outlook blue
        public static readonly Color Q4 = Color.FromRgb(0xA8, 0xA2, 0x9A);
        public static readonly Color Danger = Color.FromRgb(0xB8, 0x6B, 0x62);
        public static readonly Color Ink = Label;
        public static readonly Color InkMuted = SecondaryLabel;

        public static Brush InkBrush
        {
            get { return new SolidColorBrush(Ink); }
        }

        public static Brush MutedBrush
        {
            get { return new SolidColorBrush(InkMuted); }
        }

        public static Brush SecondaryBrush
        {
            get { return new SolidColorBrush(SecondaryLabel); }
        }

        public static Brush AccentBrush
        {
            get { return new SolidColorBrush(Accent); }
        }

        public static Brush SurfaceBrush
        {
            get { return new SolidColorBrush(Card); }
        }

        public static Brush QuadrantBrush(Core.Classification.Quadrant q)
        {
            switch (q)
            {
                case Core.Classification.Quadrant.Q1UrgentImportant:
                    return new SolidColorBrush(Q1);
                case Core.Classification.Quadrant.Q2UrgentNotImportant:
                    return new SolidColorBrush(Q2);
                case Core.Classification.Quadrant.Q3ImportantNotUrgent:
                    return new SolidColorBrush(Q3);
                default:
                    return new SolidColorBrush(Q4);
            }
        }

        public static FontFamily UiFont
        {
            get
            {
                return new FontFamily(
                    "SF Pro Text, SF Pro Display, Segoe UI Variable Text, Segoe UI, Microsoft JhengHei UI, Microsoft YaHei UI");
            }
        }

        public static UIElement AmbientField()
        {
            var grid = new System.Windows.Controls.Grid();
            var wash = new LinearGradientBrush
            {
                StartPoint = new Point(0.5, 0),
                EndPoint = new Point(0.5, 1)
            };
            wash.GradientStops.Add(new GradientStop(CanvasTop, 0.0));
            wash.GradientStops.Add(new GradientStop(Canvas, 0.55));
            wash.GradientStops.Add(new GradientStop(Color.FromRgb(0xE8, 0xE5, 0xE0), 1.0));
            grid.Background = wash;
            return grid;
        }
    }
}
