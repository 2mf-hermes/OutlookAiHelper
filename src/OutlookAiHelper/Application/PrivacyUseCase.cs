using System;
using System.IO;

namespace OutlookAiHelper.Application
{
    public sealed class PrivacyUseCase
    {
        private readonly string _root;

        public PrivacyUseCase(string rootDirectory)
        {
            _root = rootDirectory;
        }

        public string RootDirectory
        {
            get { return _root; }
        }

        public string ExportAll(string exportDirectory)
        {
            Directory.CreateDirectory(exportDirectory);
            var stamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            var target = Path.Combine(exportDirectory, "OutlookAiHelper-export-" + stamp);
            Directory.CreateDirectory(target);
            CopyIfExists(Path.Combine(_root, "settings.json"), Path.Combine(target, "settings.json"));
            CopyIfExists(Path.Combine(_root, "todos.json"), Path.Combine(target, "todos.json"));
            CopyIfExists(Path.Combine(_root, "overrides.json"), Path.Combine(target, "overrides.json"));
            return target;
        }

        public void ClearLocalData(bool clearSettings, bool clearTodos, bool clearOverrides)
        {
            if (clearSettings)
            {
                DeleteIfExists(Path.Combine(_root, "settings.json"));
            }

            if (clearTodos)
            {
                DeleteIfExists(Path.Combine(_root, "todos.json"));
            }

            if (clearOverrides)
            {
                DeleteIfExists(Path.Combine(_root, "overrides.json"));
            }

            try
            {
                var cache = Path.Combine(_root, "cache");
                if (Directory.Exists(cache))
                {
                    Directory.Delete(cache, true);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void CopyIfExists(string source, string target)
        {
            if (File.Exists(source))
            {
                File.Copy(source, target, true);
            }
        }

        private static void DeleteIfExists(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
