using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Classification;
using OutlookAiHelper.Core.Models;
using OutlookAiHelper.Localization;

namespace OutlookAiHelper.Tests
{
    public static class TodoStoreTests
    {
        public static int Run()
        {
            var failed = 0;
            var dir = Path.Combine(Path.GetTempPath(), "OutlookAiHelperTests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                var path = Path.Combine(dir, "todos.json");
                var store = new JsonTodoStore(path);
                var useCase = new TodoUseCase(store);

                failed += Check("add manual", () =>
                {
                    var item = useCase.AddManual("回覆報價");
                    return item != null && useCase.Items.Count == 1;
                });

                failed += Check("persist and reload", () =>
                {
                    var reloaded = new TodoUseCase(new JsonTodoStore(path));
                    return reloaded.Items.Count == 1 && reloaded.Items[0].Title == "回覆報價";
                });

                failed += Check("toggle done persists", () =>
                {
                    useCase.ToggleDone(useCase.Items[0].Id);
                    var reloaded = new TodoUseCase(new JsonTodoStore(path));
                    return reloaded.Items[0].Status == TodoStatus.Done;
                });

                failed += Check("remove persists", () =>
                {
                    useCase.Remove(useCase.Items[0].Id);
                    var reloaded = new TodoUseCase(new JsonTodoStore(path));
                    return reloaded.Items.Count == 0;
                });

                failed += Check("add from mail", () =>
                {
                    var mail = new MailSummary
                    {
                        EntryId = "E1",
                        Subject = "專案合約",
                        FromName = "A"
                    };
                    var item = useCase.AddFromMail(mail, "Q3");
                    return item != null && item.SourceEntryId == "E1" && item.Title == "專案合約";
                });

                failed += Check("clear all empties memory and disk", () =>
                {
                    useCase.ClearAll();
                    var reloaded = new TodoUseCase(new JsonTodoStore(path));
                    return useCase.Items.Count == 0 && reloaded.Items.Count == 0;
                });

                failed += Check("bad json quarantines", () =>
                {
                    File.WriteAllText(path, "{ not json");
                    var reloaded = new JsonTodoStore(path).Load();
                    return reloaded.Count == 0;
                });

            // Follow-up enrichment: sender snapshot, waiting time, due dates, clearing.
            var path2 = Path.Combine(dir, "todos2.json");
            var store2 = new JsonTodoStore(path2);
            var useCase2 = new TodoUseCase(store2);
            var receivedLocal = new DateTime(2026, 9, 1, 9, 30, 0, DateTimeKind.Local);

            failed += Check("added mail keeps sender and arrival time across reload", () =>
            {
                var mail = new MailSummary
                {
                    EntryId = "E2",
                    Subject = "報價單",
                    FromName = "王小明",
                    FromAddress = "ming@example.com",
                    ReceivedOn = receivedLocal
                };
                var item = useCase2.AddFromMail(mail, Strings.QuadrantName(Quadrant.Q3ImportantNotUrgent));
                if (item == null || item.SourceFrom != "王小明" || !item.SourceReceivedOn.HasValue)
                {
                    return false;
                }

                var expected = receivedLocal.ToUniversalTime();
                var reloaded = new TodoUseCase(new JsonTodoStore(path2));
                var stored = reloaded.Items[0];
                return Math.Abs((stored.SourceReceivedOn.Value - expected).TotalSeconds) < 1
                    && stored.SourceFrom == "王小明";
            });

            failed += Check("sender falls back to the address when there is no name", () =>
            {
                var mail = new MailSummary { EntryId = "E3", Subject = "S", FromAddress = "only@example.com" };
                var item = useCase2.AddFromMail(mail, null);
                return item != null && item.SourceFrom == "only@example.com";
            });

            failed += Check("waiting days come from the mail, then from creation", () =>
            {
                var now = new DateTime(2026, 10, 2, 0, 0, 0, DateTimeKind.Utc);
                var fromMail = new TodoItem
                {
                    Id = "a",
                    Status = TodoStatus.Open,
                    CreatedOn = now.AddDays(-1),
                    SourceReceivedOn = now.AddDays(-10)
                };
                var legacy = new TodoItem { Id = "b", Status = TodoStatus.Open, CreatedOn = now.AddDays(-3) };
                var clockSkew = new TodoItem { Id = "c", Status = TodoStatus.Open, CreatedOn = now.AddDays(2) };
                return fromMail.WaitingDays(now) == 10
                    && legacy.WaitingDays(now) == 3
                    && clockSkew.WaitingDays(now) == 0;
            });

            failed += Check("overdue needs an open item and a past due instant", () =>
            {
                var now = DateTime.UtcNow;
                var open = new TodoItem { Id = "a", Status = TodoStatus.Open, DueOn = now.AddHours(-1) };
                var done = new TodoItem { Id = "b", Status = TodoStatus.Done, DueOn = now.AddHours(-1) };
                var later = new TodoItem { Id = "c", Status = TodoStatus.Open, DueOn = now.AddHours(1) };
                var undated = new TodoItem { Id = "d", Status = TodoStatus.Open };
                return open.IsOverdue(now) && !done.IsOverdue(now) && !later.IsOverdue(now) && !undated.IsOverdue(now);
            });

            failed += Check("due date can be set and cleared", () =>
            {
                var id = useCase2.Items[0].Id;
                var due = new DateTime(2026, 10, 5, 23, 59, 0, DateTimeKind.Utc);
                useCase2.SetDue(id, due);
                var afterSet = new TodoUseCase(new JsonTodoStore(path2)).Items[0];
                useCase2.SetDue(id, null);
                var afterClear = new TodoUseCase(new JsonTodoStore(path2)).Items[0];
                return afterSet.DueOn.HasValue
                    && Math.Abs((afterSet.DueOn.Value - due).TotalSeconds) < 1
                    && !afterClear.DueOn.HasValue;
            });

            failed += Check("clear done removes only finished follow-ups", () =>
            {
                var before = useCase2.Items.Count;
                useCase2.ToggleDone(useCase2.Items[0].Id);
                useCase2.ToggleDone(useCase2.Items[1].Id);
                var removed = useCase2.ClearDone();
                var reloaded = new TodoUseCase(new JsonTodoStore(path2));
                return removed == 2
                    && useCase2.Items.Count == before - 2
                    && reloaded.Items.Count == before - 2
                    && reloaded.Items.All(i => i.Status == TodoStatus.Open)
                    && useCase2.ClearDone() == 0;
            });

            failed += Check("sorting: waiting longest, quadrant rank, recently updated", () =>
            {
                var now = new DateTime(2026, 10, 2, 12, 0, 0, DateTimeKind.Utc);
                var old = new TodoItem
                {
                    Id = "old",
                    Status = TodoStatus.Open,
                    CreatedOn = now.AddDays(-20),
                    UpdatedOn = now.AddDays(-20),
                    QuadrantHint = Strings.QuadrantName(Quadrant.Q3ImportantNotUrgent)
                };
                var fresh = new TodoItem
                {
                    Id = "fresh",
                    Status = TodoStatus.Open,
                    CreatedOn = now.AddDays(-1),
                    UpdatedOn = now.AddDays(-1),
                    QuadrantHint = Strings.QuadrantName(Quadrant.Q1UrgentImportant)
                };
                var finished = new TodoItem
                {
                    Id = "finished",
                    Status = TodoStatus.Done,
                    CreatedOn = now.AddDays(-30),
                    UpdatedOn = now
                };
                var all = new List<TodoItem> { fresh, finished, old };

                var byWaiting = TodoOrdering.Sort(all, TodoSortMode.LongestWaiting, now);
                var byQuadrant = TodoOrdering.Sort(all, TodoSortMode.Quadrant, now);
                var byUpdated = TodoOrdering.Sort(all, TodoSortMode.RecentlyUpdated, now);
                var manual = TodoOrdering.Sort(all, TodoSortMode.Manual, now);

                return byWaiting[0].Id == "old" && byWaiting[2].Id == "finished"
                    && byQuadrant[0].Id == "fresh" && byQuadrant[2].Id == "finished"
                    && byUpdated[0].Id == "fresh" && byUpdated[2].Id == "finished"
                    && manual[0].Id == "fresh" && manual[1].Id == "old" && manual[2].Id == "finished"
                    && TodoOrdering.QuadrantRank(Strings.QuadrantName(Quadrant.Q4Neither)) == 3
                    && TodoOrdering.QuadrantRank("something else") == TodoOrdering.UnknownQuadrantRank
                    && TodoOrdering.ParseMode("quadrant") == TodoSortMode.Quadrant
                    && TodoOrdering.ParseMode(null) == TodoSortMode.Manual
                    && TodoOrdering.ParseMode("nonsense") == TodoSortMode.Manual;
            });

            failed += Check("sorting by due date keeps undated items off the top", () =>
            {
                var now = new DateTime(2026, 10, 2, 12, 0, 0, DateTimeKind.Utc);
                var late = new TodoItem { Id = "late", Status = TodoStatus.Open, CreatedOn = now, UpdatedOn = now, DueOn = now.AddDays(-2) };
                var soon = new TodoItem { Id = "soon", Status = TodoStatus.Open, CreatedOn = now, UpdatedOn = now, DueOn = now.AddDays(1) };
                var later = new TodoItem { Id = "later", Status = TodoStatus.Open, CreatedOn = now, UpdatedOn = now, DueOn = now.AddDays(9) };
                var undated = new TodoItem { Id = "undated", Status = TodoStatus.Open, CreatedOn = now, UpdatedOn = now };
                var alreadyDone = new TodoItem { Id = "done", Status = TodoStatus.Done, CreatedOn = now, UpdatedOn = now, DueOn = now.AddDays(-5) };

                var mixed = new List<TodoItem> { undated, later, alreadyDone, soon, late };
                var sorted = TodoOrdering.Sort(mixed, TodoSortMode.DueSoonest, now);

                return sorted[0].Id == "late"
                    && sorted[1].Id == "soon"
                    && sorted[2].Id == "later"
                    && sorted[3].Id == "undated"
                    && sorted[4].Id == "done"
                    && TodoOrdering.ParseMode("DueSoonest") == TodoSortMode.DueSoonest;
            });

            failed += Check("a follow-up file from before this change still loads", () =>
            {
                var legacyPath = Path.Combine(dir, "todos-legacy.json");
                File.WriteAllText(legacyPath,
                    "{\"SchemaVersion\":1,\"Items\":[{\"Id\":\"L1\",\"Title\":\"舊資料\",\"Note\":\"\","
                    + "\"SourceEntryId\":\"E9\",\"QuadrantHint\":\"緊急重要\",\"Status\":\"Open\","
                    + "\"CreatedOn\":\"2026-01-02T03:04:05.0000000Z\",\"UpdatedOn\":\"2026-01-02T03:04:05.0000000Z\","
                    + "\"DueOn\":null}]}");
                var loaded = new JsonTodoStore(legacyPath).Load();
                return loaded.Count == 1
                    && loaded[0].Title == "舊資料"
                    && loaded[0].SourceFrom == null
                    && !loaded[0].SourceReceivedOn.HasValue
                    && loaded[0].WaitingDays(new DateTime(2026, 1, 12, 0, 0, 0, DateTimeKind.Utc)) == 9;
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
