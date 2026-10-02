using System;
using System.Collections.Generic;
using System.IO;
using OutlookAiHelper.Adapters.Storage;
using OutlookAiHelper.Application;
using OutlookAiHelper.Core.Models;

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
