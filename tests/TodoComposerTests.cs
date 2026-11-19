using System;
using System.Collections.Generic;
using System.IO;
using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// The parts of the v1.7.0 add-to-follow-up work that must be right without a window:
    /// the scan-range rule both day controls share, the date round trip behind the picker,
    /// the repeat arithmetic, the JSON round trip of the new fields, and the attachment
    /// store's copy limits. Deliberately ASCII so it compiles without a BOM.
    /// </summary>
    public static class TodoComposerTests
    {
        public static int Run()
        {
            var failed = 0;

            // --- ScanDays ---------------------------------------------------------
            failed += Check("scan range: every preset parses and is in bounds", () =>
            {
                foreach (var preset in ScanDays.Presets)
                {
                    int parsed;
                    if (!ScanDays.TryParse(preset.ToString(), out parsed) || parsed != preset)
                    {
                        return false;
                    }

                    if (preset < ScanDays.Min || preset > ScanDays.Max)
                    {
                        return false;
                    }
                }

                return ScanDays.Presets.Length == 5;
            });

            failed += Check("scan range: the bounds themselves are accepted", () =>
            {
                int low, high;
                return ScanDays.TryParse("1", out low) && low == 1
                    && ScanDays.TryParse("365", out high) && high == 365
                    && ScanDays.TryParse(" 45 ", out low) && low == 45;
            });

            failed += Check("scan range: zero, negatives, fractions, text and overflow are refused", () =>
            {
                int parsed;
                var bad = new[] { "0", "-3", "7.5", "abc", "", "   ", "366", "9999", null };
                foreach (var text in bad)
                {
                    if (ScanDays.TryParse(text, out parsed))
                    {
                        return false;
                    }
                }

                return true;
            });

            failed += Check("scan range: clamp pins to 1..365 and leaves a valid value alone", () =>
            {
                return ScanDays.Clamp(-5) == 1
                    && ScanDays.Clamp(0) == 1
                    && ScanDays.Clamp(30) == 30
                    && ScanDays.Clamp(366) == 365
                    && ScanDays.Default >= ScanDays.Min && ScanDays.Default <= ScanDays.Max;
            });

            // --- TodoDates --------------------------------------------------------
            failed += Check("due: a picked day means the end of that local day", () =>
            {
                // Relative to today, so the assertion keeps meaning "a day picked now is
                // not already overdue" no matter when the tests run.
                var day = DateTime.Today;
                var due = TodoDates.DueFromDay(day);
                if (!due.HasValue)
                {
                    return false;
                }

                // Same local calendar day back, at its last minute, and not already past.
                var local = due.Value.ToLocalTime();
                return TodoDates.DayFromDue(due) == day
                    && local.Date == day
                    && local.Hour == 23
                    && local.Minute == 59
                    && due.Value >= DateTime.UtcNow.AddMinutes(-1);
            });

            failed += Check("due: clearing the picker clears the stored value", () =>
            {
                return !TodoDates.DueFromDay(null).HasValue && !TodoDates.DayFromDue(null).HasValue;
            });

            failed += Check("due: an unspecified incoming time becomes UTC and a missing one is null", () =>
            {
                var local = new DateTime(2026, 10, 2, 9, 30, 0, DateTimeKind.Local);
                var utc = TodoDates.ToUtc(local);
                if (!utc.HasValue || utc.Value.Kind != DateTimeKind.Utc)
                {
                    return false;
                }

                if (utc.Value != local.ToUniversalTime())
                {
                    return false;
                }

                var already = new DateTime(2026, 10, 2, 1, 0, 0, DateTimeKind.Utc);
                return TodoDates.ToUtc(already) == already && !TodoDates.ToUtc(default(DateTime)).HasValue;
            });

            // --- Recurrence -------------------------------------------------------
            failed += Check("repeat: daily and weekly keep the time of day", () =>
            {
                var from = new DateTime(2026, 10, 2, 7, 15, 0, DateTimeKind.Utc);
                var nextDay = TodoRecurrence.Next(from, TodoRepeat.Daily);
                var nextWeek = TodoRecurrence.Next(from, TodoRepeat.Weekly);
                return nextDay == from.AddDays(1) && nextWeek == from.AddDays(7)
                    && nextDay.Hour == 7 && nextWeek.Kind == DateTimeKind.Utc;
            });

            failed += Check("repeat: monthly from the 31st lands on the short month's last day, then returns", () =>
            {
                var jan31 = new DateTime(2026, 1, 31, 18, 0, 0, DateTimeKind.Utc);
                var feb = TodoRecurrence.Next(jan31, TodoRepeat.Monthly, 31);
                if (feb.Month != 2 || feb.Day != DateTime.DaysInMonth(2026, 2))
                {
                    return false;
                }

                var mar = TodoRecurrence.Next(feb, TodoRepeat.Monthly, 31);
                return mar.Month == 3 && mar.Day == 31;
            });

            failed += Check("repeat: yearly from a leap day stays on the last day of February", () =>
            {
                var leap = new DateTime(2024, 2, 29, 8, 0, 0, DateTimeKind.Utc);
                var next = TodoRecurrence.Next(leap, TodoRepeat.Yearly, 29);
                return next.Year == 2025 && next.Month == 2 && next.Day == 28;
            });

            failed += Check("repeat: the stored name round-trips and an unknown name degrades to none", () =>
            {
                foreach (var mode in new[] { TodoRepeat.None, TodoRepeat.Daily, TodoRepeat.Weekly, TodoRepeat.Monthly, TodoRepeat.Yearly })
                {
                    if (TodoRecurrence.Parse(TodoRecurrence.Key(mode)) != mode)
                    {
                        return false;
                    }
                }

                return TodoRecurrence.Parse(null) == TodoRepeat.None
                    && TodoRecurrence.Parse("Fortnightly") == TodoRepeat.None
                    && !TodoRecurrence.IsRepeating(TodoRepeat.None)
                    && TodoRecurrence.IsRepeating(TodoRepeat.Weekly);
            });

            // --- Use case + JSON --------------------------------------------------
            var dir = Path.Combine(Path.GetTempPath(), "OutlookAiHelperComposer-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                var path = Path.Combine(dir, "todos.json");
                var useCase = new TodoUseCase(new JsonTodoStore(path));

                var due = TodoDates.DueFromDay(new DateTime(2026, 11, 3));
                var remind = TodoDates.DueFromDay(new DateTime(2026, 11, 2));
                var received = new DateTime(2026, 10, 1, 3, 0, 0, DateTimeKind.Utc);

                failed += Check("draft: a follow-up keeps every field through a reload", () =>
                {
                    var draft = new TodoDraft
                    {
                        Title = "Renew the contract",
                        Note = "Ask for the 2027 terms",
                        Category = "Work",
                        DueOn = due,
                        ReminderOn = remind,
                        Repeat = TodoRepeat.Monthly,
                        RepeatAnchorDay = 3,
                        SourceEntryId = "ENTRY-1",
                        SourceFrom = "Alice",
                        SourceReceivedOn = received
                    };
                    draft.NewFiles.Add(Path.Combine(dir, "ignored-in-this-test.txt"));

                    var item = useCase.AddDraft(draft, "Q3");
                    if (item == null)
                    {
                        return false;
                    }

                    var reloaded = new TodoUseCase(new JsonTodoStore(path)).Get(item.Id);
                    return reloaded != null
                        && reloaded.Title == "Renew the contract"
                        && reloaded.Note == "Ask for the 2027 terms"
                        && reloaded.Category == "Work"
                        && reloaded.DueOn == due
                        && reloaded.ReminderOn == remind
                        && reloaded.Repeat == TodoRepeat.Monthly
                        && reloaded.RepeatAnchorDay == 3
                        && reloaded.SourceEntryId == "ENTRY-1"
                        && reloaded.SourceFrom == "Alice"
                        && reloaded.SourceReceivedOn == received
                        && reloaded.QuadrantHint == "Q3";
                });

                failed += Check("draft: an empty title is refused instead of saving a blank row", () =>
                {
                    var before = useCase.Items.Count;
                    var empty = useCase.AddDraft(new TodoDraft { Title = "   " }, "Q4");
                    return empty == null && useCase.Items.Count == before;
                });

                failed += Check("edit: the title, note and attachment list are replaced", () =>
                {
                    var item = useCase.Items[0];
                    var kept = new List<TodoAttachment>
                    {
                        TodoAttachment.Create("brief.pdf", "brief-1.pdf", 1234)
                    };
                    useCase.UpdateDraft(item.Id, new TodoDraft
                    {
                        Title = "Renew the contract (revised)",
                        Note = "New terms",
                        Category = "Legal",
                        DueOn = null,
                        Repeat = TodoRepeat.None
                    }, kept);

                    var reloaded = new TodoUseCase(new JsonTodoStore(path)).Get(item.Id);
                    return reloaded != null
                        && reloaded.Title == "Renew the contract (revised)"
                        && reloaded.Note == "New terms"
                        && reloaded.Category == "Legal"
                        && !reloaded.DueOn.HasValue
                        && reloaded.Repeat == TodoRepeat.None
                        && reloaded.Attachments.Count == 1
                        && reloaded.Attachments[0].StoredName == "brief-1.pdf"
                        && reloaded.Attachments[0].SizeBytes == 1234;
                });

                failed += Check("repeat: ticking off a repeating follow-up moves the deadline and keeps it open", () =>
                {
                    var draft = new TodoDraft
                    {
                        Title = "Water the plants",
                        DueOn = due,
                        ReminderOn = remind,
                        Repeat = TodoRepeat.Monthly,
                        RepeatAnchorDay = 3
                    };
                    var item = useCase.AddDraft(draft, "Q4");
                    var next = useCase.ToggleDone(item.Id);
                    var after = useCase.Get(item.Id);
                    return next.HasValue
                        && after.Status == TodoStatus.Open
                        && after.DueOn == next
                        && after.ReminderOn.HasValue
                        && after.ReminderOn.Value - remind.Value == next.Value - due.Value;
                });

                failed += Check("compat: a file written before these fields still loads", () =>
                {
                    var legacy = Path.Combine(dir, "legacy.json");
                    File.WriteAllText(legacy,
                        "{\"SchemaVersion\":1,\"Items\":[{\"Id\":\"old-1\",\"Title\":\"Old item\",\"Status\":\"Open\"}]}");
                    var loaded = new JsonTodoStore(legacy).Load();
                    return loaded.Count == 1
                        && loaded[0].Title == "Old item"
                        && loaded[0].Repeat == TodoRepeat.None
                        && !loaded[0].DueOn.HasValue
                        && (loaded[0].Attachments == null || loaded[0].Attachments.Count == 0);
                });

                // --- Attachments --------------------------------------------------
                failed += Check("attachments: a copied file keeps its bytes and gets a safe stored name", () =>
                {
                    var store = new TodoAttachmentStore(Path.Combine(dir, "attach"));
                    var source = Path.Combine(dir, "report.txt");
                    File.WriteAllText(source, "hello attachment");

                    IList<string> errors;
                    var added = store.Import("todo-1", new[] { source }, out errors);
                    if (added.Count != 1 || errors.Count != 0)
                    {
                        return false;
                    }

                    var copy = store.PathOf("todo-1", added[0]);
                    return File.Exists(copy)
                        && File.ReadAllText(copy) == "hello attachment"
                        && added[0].FileName == "report.txt"
                        && added[0].StoredName.IndexOf("..", StringComparison.Ordinal) < 0
                        && added[0].SizeBytes == new FileInfo(source).Length;
                });

                failed += Check("attachments: an oversized file is reported and not attached", () =>
                {
                    var store = new TodoAttachmentStore(Path.Combine(dir, "small"), 4);
                    var source = Path.Combine(dir, "big.txt");
                    File.WriteAllText(source, "too large");

                    IList<string> errors;
                    var added = store.Import("todo-2", new[] { source }, out errors);
                    return added.Count == 0 && errors.Count == 1;
                });

                failed += Check("attachments: removing an attachment deletes its copy", () =>
                {
                    var store = new TodoAttachmentStore(Path.Combine(dir, "attach"));
                    var source = Path.Combine(dir, "second.txt");
                    File.WriteAllText(source, "second");

                    IList<string> errors;
                    var added = store.Import("todo-3", new[] { source }, out errors);
                    var copy = store.PathOf("todo-3", added[0]);
                    store.Remove("todo-3", added[0].StoredName);
                    return !File.Exists(copy);
                });
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch (Exception) { }
            }

            return failed;
        }

        private static int Check(string name, Func<bool> body)
        {
            try
            {
                if (body())
                {
                    Console.WriteLine("PASS " + name);
                    return 0;
                }

                Console.WriteLine("FAIL " + name);
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL " + name + " :: " + ex.Message);
                return 1;
            }
        }
    }
}
