using System;
using System.Collections.Generic;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Tests
{
    /// <summary>
    /// Scan scope must never pick up deleted or junk mail. These tests cover the
    /// pure decision logic used by the Outlook reader's folder walk.
    /// </summary>
    public static class ScanScopeTests
    {
        public static int Run()
        {
            var failed = 0;

            failed += Check("deleted folder name excluded in every UI language", () =>
                ScanFolderFilter.IsDeletedFolderName("Deleted Items")
                && ScanFolderFilter.IsDeletedFolderName("刪除的郵件")
                && ScanFolderFilter.IsDeletedFolderName("已删除邮件")
                && ScanFolderFilter.IsDeletedFolderName("已删除的邮件")
                && ScanFolderFilter.IsDeletedFolderName("  deleted items  "));

            failed += Check("mail folders are kept", () =>
                !ScanFolderFilter.IsDeletedFolderName("Inbox")
                && !ScanFolderFilter.IsDeletedFolderName("收件匣")
                && !ScanFolderFilter.IsDeletedFolderName("專案/待處理")
                && !ScanFolderFilter.IsDeletedFolderName("Deleted Items Archive")
                && !ScanFolderFilter.IsDeletedFolderName(null)
                && !ScanFolderFilter.IsDeletedFolderName(""));

            failed += Check("deleted folder excluded by EntryID even when renamed", () =>
            {
                var ids = new List<string> { "AAAA-DELETED-ID" };
                return ScanFolderFilter.IsExcluded("我的舊信", "AAAA-DELETED-ID", ids)
                    && ScanFolderFilter.IsExcluded("我的舊信", "aaaa-deleted-id", ids)
                    && !ScanFolderFilter.IsExcluded("我的舊信", "OTHER-ID", ids);
            });

            failed += Check("excluded by name without any EntryID map", () =>
            {
                var ids = new List<string>();
                return ScanFolderFilter.IsExcluded("刪除的郵件", null, ids)
                    && ScanFolderFilter.IsExcluded("Deleted Items", null, null)
                    && !ScanFolderFilter.IsExcluded("Inbox", null, null);
            });

            failed += Check("normal folder with unknown id stays in scope", () =>
                !ScanFolderFilter.IsExcluded("Inbox", "INBOX-ID", new List<string> { "DELETED-ID" })
                && !ScanFolderFilter.IsExcluded(null, null, null));

            failed += Check("junk folder name excluded in every UI language", () =>
                ScanFolderFilter.IsJunkFolderName("Junk E-mail")
                && ScanFolderFilter.IsJunkFolderName("Junk Email")
                && ScanFolderFilter.IsJunkFolderName("junk e-mail")
                && ScanFolderFilter.IsJunkFolderName("  垃圾郵件  ")
                && ScanFolderFilter.IsJunkFolderName("垃圾邮件")
                && !ScanFolderFilter.IsJunkFolderName("Junk E-mail Archive")
                && !ScanFolderFilter.IsJunkFolderName("Inbox")
                && !ScanFolderFilter.IsJunkFolderName(null)
                && !ScanFolderFilter.IsJunkFolderName(""));

            failed += Check("junk folder excluded by EntryID even when renamed", () =>
            {
                var ids = new List<string> { "AAAA-JUNK-ID" };
                return ScanFolderFilter.IsExcluded("廣告信", "AAAA-JUNK-ID", ids)
                    && ScanFolderFilter.IsExcluded("廣告信", "aaaa-junk-id", ids)
                    && !ScanFolderFilter.IsExcluded("廣告信", "OTHER-ID", ids);
            });

            failed += Check("junk excluded by name without any EntryID map", () =>
            {
                var ids = new List<string>();
                return ScanFolderFilter.IsExcluded("Junk E-mail", null, ids)
                    && ScanFolderFilter.IsExcluded("垃圾郵件", null, null)
                    && ScanFolderFilter.IsExcludedFolderName("Junk Email")
                    && ScanFolderFilter.IsExcludedFolderName("Deleted Items")
                    && !ScanFolderFilter.IsExcludedFolderName("Junk");
            });

            failed += Check("Outlook folder type constants match OlDefaultFolders", () =>
                ScanFolderFilter.OlFolderDeletedItems == 3 && ScanFolderFilter.OlFolderJunk == 23);

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
