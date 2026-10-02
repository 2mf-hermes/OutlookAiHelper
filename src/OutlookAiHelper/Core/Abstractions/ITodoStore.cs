using System.Collections.Generic;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Core.Abstractions
{
    public interface ITodoStore
    {
        IList<TodoItem> Load();
        void Save(IList<TodoItem> items);
    }
}
