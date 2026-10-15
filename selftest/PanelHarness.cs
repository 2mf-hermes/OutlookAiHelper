// Panel harness: opens the real window in-process, drives the panel's entry points without a mouse,
// then reads the rendered panel back as text plus a PNG.
//
// Why in-process: the panel column is collapsed until a row is clicked, and a synthetic click needs
// the window in the foreground - which is exactly what the machine refuses to give an unattended
// process. Calling the window's own handlers on the UI thread needs no mouse and no focus.
//
// Usage: panelharness.exe <outTxt> [outPng] [dataFile]
//   dataFile: key=value fixture (see fixture.txt next to it) plus assertions, one set per pass:
//     expect= / absent=   pass 1: a follow-up is selected, its mail is in the scan
//     expect2= / absent2= pass 2: a follow-up is selected, its mail is out of the scan
//     expect3= / absent3= pass 3: a mail row is selected from the list, no follow-up behind it
//     expect4= / absent4= pass 4: a follow-up is selected, its topic holds no other mail
//     expect5= / absent5= pass 5: a follow-up is selected and one of its related mails is clicked
//   Only text that is actually on screen counts: an element inside a collapsed card is in the tree
//   with a zero size, and used to read as a pass by mistake.
//
// Keep display text out of this file: it comes from fixture.txt, because a .cs compiled without a
// BOM is read in the console codepage and would garble CJK literals.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Classification;
using OutlookAiHelper.Core.Models;

internal static class PanelHarness
{
    private static readonly StringBuilder Log = new StringBuilder();
    private static string _outTxt;

    private static readonly string[] FieldNames = new string[]
    {
        "_detailScroll", "_detailTitle", "_mailFactsBlock", "_mailFactsRows", "_reasonPanel",
        "_relatedBlock", "_todoActions", "_openOutlookButton"
    };

    [STAThread]
    private static int Main(string[] args)
    {
        _outTxt = args.Length > 0 ? args[0] : "panel-harness.txt";
        string outPng = args.Length > 1 && args[1].Length > 0 ? args[1] : null;
        string dataFile = args.Length > 2 && args[2].Length > 0 ? args[2] : null;
        int exit = 0;

        try
        {
            Application app = new Application();
            app.ShutdownMode = ShutdownMode.OnExplicitShutdown;

            // The harness sits beside a copy of the app, so the assembly resolves by name alone - one
            // instance only, so the TodoItem handed to the window is the window's own type.
            Type type = Type.GetType("OutlookAiHelper.UI.MainWindow, OutlookAiHelper", false);
            if (type == null)
            {
                Say("FATAL: OutlookAiHelper.UI.MainWindow not found - is a copy of the app beside this exe?");
                return 4;
            }

            Window win = (Window)Activator.CreateInstance(type);
            win.WindowStartupLocation = WindowStartupLocation.Manual;
            win.Left = 40;
            win.Top = 40;
            win.Width = 1180;
            win.Height = 760;
            win.Show();
            Pump();
            win.UpdateLayout();

            Say("window: " + win.Title);
            Say(string.Format(CultureInfo.InvariantCulture, "window: {0}x{1} at {2},{3}",
                (int)win.ActualWidth, (int)win.ActualHeight, (int)win.Left, (int)win.Top));

            FrameworkElement host = Field(win, "_detailHost") as FrameworkElement;
            if (host == null)
            {
                Say("FATAL: _detailHost is not reachable by reflection.");
                return 3;
            }

            IList items = Field(win, "_items") as IList;
            if (items == null)
            {
                Say("FATAL: _items is not reachable by reflection.");
                return 3;
            }

            Say("_detailHost before any selection: " + host.Visibility);
            Say("");

            string dir = dataFile == null ? null : Path.GetDirectoryName(Path.GetFullPath(dataFile));
            Dictionary<string, string> fx = ReadPairs(dir == null ? null : Path.Combine(dir, "fixture.txt"),
                dataFile);

            string entryId = Fx(fx, "entryId", "smoke-entry-1");
            string title = Fx(fx, "title", "follow-up");
            string note = Fx(fx, "note", "probe");
            string sender = Fx(fx, "sender", "sender@example.com");
            string address = Fx(fx, "address", "sender@example.com");
            string folder = Fx(fx, "folder", "Inbox");
            string quadrant = Fx(fx, "quadrant", "Q1");
            string subject = Fx(fx, "subject", title);
            string replyPrefix = Fx(fx, "replyPrefix", "RE: ");
            string reasonDetail = Fx(fx, "reasonDetail", "");

            TodoItem todo = TodoItem.Create(title, entryId, quadrant, sender, DateTime.Now.AddDays(-2));
            todo.Note = note;

            var source = MakeMail(entryId, subject, sender, address, folder, reasonDetail);
            var reply = MakeMail(entryId + "-reply", replyPrefix + subject, sender, address, folder, reasonDetail);

            exit |= Pass(win, host, items, todo, "ShowTodoDetail", new ScanResultItem[] { source, reply },
                "the follow-up's mail is in the scan (2 mails of one topic)", fx, "expect", "absent", 1,
                outPng);
            exit |= Pass(win, host, items, todo, "ShowTodoDetail", new ScanResultItem[0],
                "the follow-up's mail is out of the scan", fx, "expect2", "absent2", 2, outPng);
            // The other way into the panel: a mail row clicked in the list, with no follow-up
            // behind it. A synthetic click cannot be raised here, so the renderer the row's own
            // handler calls is invoked directly - MainWindow.cs:2363 and :2387.
            exit |= Pass(win, host, items, source, "OnMailSelectedDetail",
                new ScanResultItem[] { source, reply },
                "a mail row is selected from the list, with no follow-up", fx, "expect3", "absent3", 3,
                outPng);
            // A follow-up whose topic holds only its own mail: the topic card must stay down rather
            // than stand there empty, and the follow-up's own cards stay up.
            exit |= Pass(win, host, items, todo, "ShowTodoDetail", new ScanResultItem[] { source },
                "the follow-up's topic holds no other mail", fx, "expect4", "absent4", 4, outPng);

            // A click on a related mail walks back to the follow-up through the window's own store,
            // so the harness's follow-up has to be in it (in memory only - nothing is saved).
            IList todos = Todos(win);
            if (todos != null && !todos.Contains(todo))
            {
                todos.Add(todo);
            }

            // The related mail of the topic, clicked the way its row's handler reaches it: that
            // mail's own facts go into the card above, and the card's own open action comes back.
            exit |= Pass(win, host, items, todo, "ShowTodoDetail", new ScanResultItem[] { source, reply },
                "a related mail of the topic is clicked", fx, "expect5", "absent5", 5, outPng,
                "OnRelatedMailClick", reply);

            win.Close();
            app.Shutdown();
        }
        catch (Exception ex)
        {
            Say("HARNESS ERROR: " + ex.ToString());
            exit = 2;
        }

        File.WriteAllText(_outTxt, Log.ToString(), new UTF8Encoding(false));
        Console.WriteLine(Log.ToString());
        return exit;
    }

    // ------------------------------------------------------------------ one rendered state

    private static int Pass(Window win, FrameworkElement host, IList items, object subject,
        string entry, ScanResultItem[] mails, string caption, Dictionary<string, string> fx,
        string expectKey, string absentKey, int index, string pngBase,
        string afterEntry = null, object afterArg = null)
    {
        int exit = 0;
        Say("=== pass " + index + ": " + caption + " ===");

        items.Clear();
        if (mails != null)
        {
            foreach (ScanResultItem m in mails)
            {
                items.Add(m);
            }
        }

        Say("  the scan holds " + items.Count + " mail(s)");
        // The window's own entry point drives each pass: the follow-up selection handler for the
        // follow-up paths, and the renderer the mail-row click reaches for the mail-list path.
        if (!Invoke(win, entry, subject))
        {
            Say("  FATAL: the panel's entry point " + entry + "() is not reachable by reflection.");
            return 4;
        }

        if (afterEntry != null)
        {
            Say("  then " + afterEntry + "() with the related mail, the way a click on its row reaches it");
            if (!Invoke(win, afterEntry, afterArg))
            {
                Say("  FATAL: the panel's entry point " + afterEntry + "() is not reachable by reflection.");
                return 4;
            }
        }

        Pump();
        win.UpdateLayout();

        Say(string.Format(CultureInfo.InvariantCulture, "  _detailHost: {0} {1}x{2}",
            host.Visibility, (int)host.ActualWidth, (int)host.ActualHeight));
        Say("  --- the panel's parts ---");
        foreach (string name in FieldNames)
        {
            Say("  " + Describe(win, name));
        }

        StackPanel stack = PanelStack(win);
        Say("  --- the panel, block by block, top to bottom ---");
        if (stack == null)
        {
            Say("    (the panel's stack is not reachable)");
        }
        else
        {
            foreach (UIElement child in stack.Children)
            {
                FrameworkElement fe = child as FrameworkElement;
                if (fe == null)
                {
                    continue;
                }

                Point p = PointOf(fe, stack);
                bool onScreen = OnScreen(fe);
                Say(string.Format(CultureInfo.InvariantCulture,
                    "    {0,-9} y={1,5} {2,4}x{3,-4}  {4}",
                    onScreen ? "visible" : "collapsed", (int)p.Y, (int)fe.ActualWidth,
                    (int)fe.ActualHeight, FirstText(fe)));
            }
        }

        // Only what is really on screen counts: a collapsed card's text still walks as a TextBlock.
        HashSet<string> visible = new HashSet<string>(StringComparer.Ordinal);
        List<string> onScreenRows = new List<string>();
        List<string> hiddenRows = new List<string>();
        List<string> cards = new List<string>();
        foreach (DependencyObject d in Descendants(host))
        {
            FrameworkElement fe = d as FrameworkElement;
            if (fe == null)
            {
                continue;
            }

            string text = TextOf(d);
            if (text != null)
            {
                string trimmed = text.Trim();
                if (trimmed.Length > 0)
                {
                    Point p = PointOf(fe, host);
                    string row = string.Format(CultureInfo.InvariantCulture, "    y={0,5} x={1,4} {2,4}x{3,-4}  {4}",
                        (int)p.Y, (int)p.X, (int)fe.ActualWidth, (int)fe.ActualHeight, trimmed);
                    if (OnScreen(fe))
                    {
                        visible.Add(trimmed);
                        onScreenRows.Add(row);
                    }
                    else
                    {
                        hiddenRows.Add(row);
                    }
                }
            }

            Border border = d as Border;
            if (border != null && fe.ActualWidth >= 180 && fe.ActualHeight >= 24)
            {
                Point p = PointOf(fe, host);
                cards.Add(string.Format(CultureInfo.InvariantCulture,
                    "    {0,-9} y={1,5} {2,4}x{3,-4} radius={4}",
                    OnScreen(fe) ? "visible" : "collapsed", (int)p.Y, (int)fe.ActualWidth,
                    (int)fe.ActualHeight, (int)border.CornerRadius.TopLeft));
            }
        }

        onScreenRows.Sort(StringComparer.Ordinal);
        hiddenRows.Sort(StringComparer.Ordinal);
        cards.Sort(StringComparer.Ordinal);

        Say("  --- cards (a border wide enough to be one) ---");
        foreach (string c in cards) Say(c);
        Say("  --- labels on screen (" + onScreenRows.Count + ") ---");
        foreach (string r in onScreenRows) Say(r);
        if (hiddenRows.Count > 0)
        {
            Say("  --- labels in the tree but not on screen (" + hiddenRows.Count + ") ---");
            foreach (string r in hiddenRows) Say(r);
        }

        if (pngBase != null)
        {
            WritePng(win, host, pngBase, index);
        }

        Say("");

        var clicked = afterArg as ScanResultItem;
        if (clicked != null)
        {
            exit |= CheckOpenTarget(win, clicked.Mail.EntryId);
        }

        exit |= Assert(fx, expectKey, visible, true);
        exit |= Assert(fx, absentKey, visible, false);
        Say("");
        return exit;
    }

    private static int Assert(Dictionary<string, string> fx, string key, HashSet<string> visible, bool wanted)
    {
        // The assertion keys repeat by design, so they are read from the joined list, not the map.
        string blob = null;
        if (!fx.TryGetValue("#" + key, out blob) || blob == null || blob.Length == 0)
        {
            return 0;
        }

        string[] wanted2 = blob.Split('\n');
        int ok = 0;
        Say("  --- " + key + " (" + (wanted ? "must be on screen" : "must not be on screen") + ") ---");
        foreach (string w0 in wanted2)
        {
            string w = w0.Trim();
            if (w.Length == 0)
            {
                continue;
            }

            bool hit = visible.Contains(w);
            if (hit == wanted) ok++;
            else if (wanted) Say("    MISSING: '" + w + "' is not visible");
            else Say("    UNEXPECTED: '" + w + "' is visible");
        }

        Say("  " + key + ": " + ok + "/" + wanted2.Length);
        return ok == wanted2.Length ? 0 : 1;
    }

    /// <summary>
    /// Text alone cannot prove the card's open action points at the related mail the click
    /// picked -- the follow-up card carries its own open button -- so the target the handler
    /// reads is checked here, plus that the button is really on screen and clickable.
    /// </summary>
    private static int CheckOpenTarget(Window win, string wantedEntryId)
    {
        int ok = 0;
        int total = 3;

        string target = Field(win, "_mailFactsEntryId") as string;
        if (string.Equals(target, wantedEntryId, StringComparison.OrdinalIgnoreCase)) ok++;
        else Say("    TARGET: the card's open action points at '" + (target ?? "(none)") + "', not '" + wantedEntryId + "'");

        Button button = Field(win, "_openOutlookButton") as Button;
        if (button != null && button.IsEnabled) ok++;
        else Say("    TARGET: the card's open button is not enabled");

        if (button != null && OnScreen(button)) ok++;
        else Say("    TARGET: the card's open button is not on screen");

        Say("  mailOpenTarget: " + ok + "/" + total);
        return ok == total ? 0 : 1;
    }

    private static IList Todos(Window win)
    {
        object useCase = Field(win, "_todos");
        if (useCase == null)
        {
            return null;
        }

        PropertyInfo p = useCase.GetType().GetProperty("Items", BindingFlags.Instance | BindingFlags.Public);
        return p == null ? null : p.GetValue(useCase, null) as IList;
    }

    // ------------------------------------------------------------------ reflection

    private static bool Invoke(object target, string method, object arg)
    {
        MethodInfo m = target.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic);
        if (m == null)
        {
            return false;
        }

        m.Invoke(target, new object[] { arg });
        return true;
    }

    private static object Field(object target, string name)
    {
        FieldInfo f = target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
        return f == null ? null : f.GetValue(target);
    }

    private static string Describe(object target, string name)
    {
        FrameworkElement fe = Field(target, name) as FrameworkElement;
        if (fe == null)
        {
            return name + ": (not there)";
        }

        return string.Format(CultureInfo.InvariantCulture, "{0}: {1} {2}x{3}",
            name, fe.Visibility, (int)fe.ActualWidth, (int)fe.ActualHeight);
    }

    // ------------------------------------------------------------------ visual tree

    private static StackPanel PanelStack(Window win)
    {
        // The panel's stack is what the scroller holds, since the card layout moved inside it.
        ScrollViewer sc = Field(win, "_detailScroll") as ScrollViewer;
        if (sc != null && sc.Content is StackPanel)
        {
            return (StackPanel)sc.Content;
        }

        FrameworkElement host = Field(win, "_detailHost") as FrameworkElement;
        ScrollViewer found = FindScroller(host);
        return found == null ? null : found.Content as StackPanel;
    }

    private static ScrollViewer FindScroller(DependencyObject root)
    {
        foreach (DependencyObject d in Descendants(root))
        {
            ScrollViewer sv = d as ScrollViewer;
            if (sv != null && sv.ActualWidth >= 180)
            {
                return sv;
            }
        }

        return null;
    }

    private static IEnumerable<DependencyObject> Descendants(DependencyObject root)
    {
        var stack = new Stack<DependencyObject>();
        stack.Push(root);
        while (stack.Count > 0)
        {
            DependencyObject node = stack.Pop();
            yield return node;
            int n = 0;
            try { n = VisualTreeHelper.GetChildrenCount(node); }
            catch (Exception) { n = 0; }
            for (int i = 0; i < n; i++)
            {
                DependencyObject child = null;
                try { child = VisualTreeHelper.GetChild(node, i); }
                catch (Exception) { child = null; }
                if (child != null)
                {
                    stack.Push(child);
                }
            }
        }
    }

    private static bool OnScreen(FrameworkElement fe)
    {
        return fe.IsVisible && fe.ActualWidth > 0.5 && fe.ActualHeight > 0.5;
    }

    private static Point PointOf(FrameworkElement fe, Visual ancestor)
    {
        try
        {
            return fe.TransformToAncestor(ancestor).Transform(new Point(0, 0));
        }
        catch (Exception)
        {
            return new Point(-1, -1);
        }
    }

    private static string TextOf(DependencyObject d)
    {
        TextBlock tb = d as TextBlock;
        if (tb != null)
        {
            return tb.Text;
        }

        TextBox box = d as TextBox;
        if (box != null)
        {
            return box.Text;
        }

        ContentControl cc = d as ContentControl;
        if (cc != null)
        {
            return cc.Content as string;
        }

        return null;
    }

    private static string FirstText(DependencyObject root)
    {
        foreach (DependencyObject d in Descendants(root))
        {
            string t = TextOf(d);
            if (t != null && t.Trim().Length > 0)
            {
                return t.Trim();
            }
        }

        return "(no text)";
    }

    private static void WritePng(Window win, FrameworkElement host, string baseName, int index)
    {
        try
        {
            ScrollViewer sc = FindScroller(host);
            FrameworkElement target = host;
            if (sc != null && sc.Content is FrameworkElement)
            {
                target = (FrameworkElement)sc.Content;
            }

            double w = target.ActualWidth > 1 ? target.ActualWidth : host.ActualWidth;
            double h = target.ActualHeight > 1 ? target.ActualHeight : host.ActualHeight;
            if (h > 4000) h = 4000;
            var rtb = new RenderTargetBitmap((int)Math.Ceiling(w), (int)Math.Ceiling(h), 96, 96,
                PixelFormats.Pbgra32);
            rtb.Render(target);
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(rtb));
            string path = Path.Combine(Path.GetDirectoryName(baseName),
                Path.GetFileNameWithoutExtension(baseName) + "-" + index + Path.GetExtension(baseName));
            using (FileStream fs = File.Create(path))
            {
                encoder.Save(fs);
            }

            Say(string.Format(CultureInfo.InvariantCulture, "  png: {0} ({1}x{2}, from {3})",
                path, (int)Math.Ceiling(w), (int)Math.Ceiling(h), target.GetType().Name));
        }
        catch (Exception ex)
        {
            Say("  png failed: " + ex.Message);
        }
    }

    // ------------------------------------------------------------------ fixture text

    private static string Fx(Dictionary<string, string> fx, string key, string fallback)
    {
        string v = null;
        if (fx.TryGetValue(key, out v) && v != null && v.Length > 0)
        {
            return v;
        }

        return fallback;
    }

    /// <summary>Reads key=value lines; every line is kept, so a key may repeat (the assertions do).</summary>
    private static Dictionary<string, string> ReadPairs(string fixturePath, string extraPath)
    {
        var fx = new List<KeyValuePair<string, string>>();
        ReadInto(fixturePath, fx);
        ReadInto(extraPath, fx);

        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, string> kv in fx)
        {
            if (!map.ContainsKey(kv.Key))
            {
                map[kv.Key] = kv.Value;
            }
        }

        // The assertion keys repeat by design; the caller reads them from the joined lists below.
        foreach (string k in new string[] { "expect", "absent", "expect2", "absent2", "expect3", "absent3", "expect4", "absent4", "expect5", "absent5" })
        {
            var vals = new List<string>();
            foreach (KeyValuePair<string, string> kv in fx)
            {
                if (kv.Key == k) vals.Add(kv.Value);
            }

            map["#" + k] = string.Join("\n", vals.ToArray());
        }

        return map;
    }

    private static void ReadInto(string path, List<KeyValuePair<string, string>> into)
    {
        if (path == null || !File.Exists(path))
        {
            return;
        }

        foreach (string raw in File.ReadAllLines(path, Encoding.UTF8))
        {
            string line = raw == null ? string.Empty : raw.Trim();
            if (line.Length == 0 || line.StartsWith("#"))
            {
                continue;
            }

            int eq = line.IndexOf('=');
            if (eq <= 0)
            {
                continue;
            }

            into.Add(new KeyValuePair<string, string>(line.Substring(0, eq).Trim(),
                line.Substring(eq + 1).Trim()));
        }
    }

    private static ScanResultItem MakeMail(string entryId, string subject, string sender, string address,
        string folder, string reasonDetail)
    {
        var mail = new MailSummary();
        mail.EntryId = entryId;
        mail.Subject = subject;
        mail.FromName = sender;
        mail.FromAddress = address;
        mail.FolderPath = folder;
        mail.ReceivedOn = DateTime.Now.AddDays(-2).AddHours(-3);
        mail.IsRead = false;
        mail.HasFlag = false;
        mail.IsMeetingRequest = true;
        mail.Importance = 0;

        var reasons = new List<ScoreReason>();
        reasons.Add(new ScoreReason("fixture", "reason.signal.importance", 3, reasonDetail));

        var item = new ScanResultItem();
        item.Mail = mail;
        item.Classification = new ClassificationResult(entryId, Quadrant.Q1UrgentImportant, 8, 6,
            reasons, false);
        return item;
    }

    // ------------------------------------------------------------------ plumbing

    private static void Pump()
    {
        for (int i = 0; i < 4; i++)
        {
            Dispatcher.CurrentDispatcher.Invoke(new Action(delegate { }), DispatcherPriority.Background);
            System.Threading.Thread.Sleep(60);
        }
    }

    private static void Say(string line)
    {
        Log.AppendLine(line);
    }
}
