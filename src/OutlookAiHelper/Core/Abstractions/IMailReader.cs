using System;
using System.Collections.Generic;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Core.Abstractions
{
    public sealed class MailScanScope
    {
        public IList<string> FolderPaths { get; set; }
        public int Days { get; set; }
    }

    public interface IMailReader
    {
        bool IsAvailable(out string reasonKey);
        IList<string> ListFolders();
        IEnumerable<MailSummary> Enumerate(MailScanScope scope, Action<int, int> progress, Func<bool> isCancelled);
    }
}
