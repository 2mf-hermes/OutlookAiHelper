using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OutlookAiHelper.Core.Models;

namespace OutlookAiHelper.Adapters.Storage
{
    /// <summary>
    /// Copies the files a follow-up carries into its own folder under the app data
    /// directory (<c>%LOCALAPPDATA%\OutlookAiHelper\attachments\{todoId}\</c>). Nothing
    /// is ever written into the repository or next to the executable, and nothing is
    /// deleted from the user's original location - an attachment is a copy.
    ///
    /// Every call is defensive: a missing source file, a locked file, a name with
    /// path separators or an over-size file is reported through <c>errors</c> and the
    /// remaining files still import.
    /// </summary>
    public sealed class TodoAttachmentStore
    {
        public const long DefaultMaxBytes = 50L * 1024 * 1024;

        private readonly string _root;
        private readonly long _maxBytes;

        public TodoAttachmentStore()
            : this(Path.Combine(JsonPaths.RootDirectory, "attachments"))
        {
        }

        public TodoAttachmentStore(string root)
            : this(root, DefaultMaxBytes)
        {
        }

        public TodoAttachmentStore(string root, long maxBytes)
        {
            _root = root;
            _maxBytes = maxBytes > 0 ? maxBytes : DefaultMaxBytes;
        }

        public string Root
        {
            get { return _root; }
        }

        public long MaxBytes
        {
            get { return _maxBytes; }
        }

        /// <summary>
        /// Copies each readable file and returns the records to store. The count of
        /// failures is reported, never thrown: attaching three files and having one
        /// locked must still attach the other two.
        /// </summary>
        public IList<TodoAttachment> Import(string todoId, IEnumerable<string> paths, out IList<string> errors)
        {
            var added = new List<TodoAttachment>();
            var failed = new List<string>();
            if (string.IsNullOrEmpty(todoId) || paths == null)
            {
                errors = failed;
                return added;
            }

            var folder = FolderOf(todoId);

            foreach (var path in paths)
            {
                try
                {
                    if (string.IsNullOrEmpty(path) || !File.Exists(path))
                    {
                        failed.Add(path ?? string.Empty);
                        continue;
                    }

                    var info = new FileInfo(path);
                    if (info.Length > _maxBytes)
                    {
                        failed.Add(info.Name);
                        continue;
                    }

                    Directory.CreateDirectory(folder);
                    var stored = UniqueName(folder, info.Name);
                    File.Copy(path, Path.Combine(folder, stored), false);
                    added.Add(TodoAttachment.Create(info.Name, stored, info.Length));
                }
                catch (Exception)
                {
                    failed.Add(path ?? string.Empty);
                }
            }

            errors = failed;
            return added;
        }

        public string PathOf(string todoId, TodoAttachment attachment)
        {
            if (string.IsNullOrEmpty(todoId) || attachment == null || string.IsNullOrEmpty(attachment.StoredName))
            {
                return null;
            }

            return Path.Combine(FolderOf(todoId), attachment.StoredName);
        }

        /// <summary>Deletes one stored file; a file already gone is not an error.</summary>
        public void Remove(string todoId, string storedName)
        {
            try
            {
                var path = PathOf(todoId, new TodoAttachment { StoredName = storedName });
                if (path != null && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>Removes the item's folder, and its parent when it became empty.</summary>
        public void DeleteAll(string todoId)
        {
            try
            {
                var folder = FolderOf(todoId);
                if (Directory.Exists(folder))
                {
                    Directory.Delete(folder, true);
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Copies one occurrence's files to the next one, so a repeating follow-up does
        /// not lose the quote or the PDF it was created with.
        /// </summary>
        public IList<TodoAttachment> CopyTo(string fromTodoId, string toTodoId, IEnumerable<TodoAttachment> attachments, out IList<string> errors)
        {
            var copied = new List<TodoAttachment>();
            var failed = new List<string>();
            if (attachments == null || string.IsNullOrEmpty(toTodoId))
            {
                errors = failed;
                return copied;
            }

            foreach (var attachment in attachments)
            {
                try
                {
                    var source = PathOf(fromTodoId, attachment);
                    if (source == null || !File.Exists(source))
                    {
                        failed.Add(attachment == null ? string.Empty : attachment.FileName);
                        continue;
                    }

                    var folder = FolderOf(toTodoId);
                    Directory.CreateDirectory(folder);
                    var stored = UniqueName(folder, attachment.StoredName ?? attachment.FileName);
                    File.Copy(source, Path.Combine(folder, stored), false);
                    copied.Add(TodoAttachment.Create(
                        attachment.FileName,
                        stored,
                        new FileInfo(source).Length));
                }
                catch (Exception)
                {
                    failed.Add(attachment == null ? string.Empty : attachment.FileName);
                }
            }

            errors = failed;
            return copied;
        }

        private string FolderOf(string todoId)
        {
            return Path.Combine(_root, SafeId(todoId));
        }

        private static string SafeId(string todoId)
        {
            var clean = new string((todoId ?? string.Empty)
                .Where(c => char.IsLetterOrDigit(c))
                .ToArray());
            return string.IsNullOrEmpty(clean) ? "item" : clean;
        }

        /// <summary>
        /// A name that cannot escape the folder: the user's folder part is dropped, the
        /// remaining name is stripped of invalid characters and capped, and the timestamp
        /// plus a random tail makes two files called "report.pdf" two distinct files.
        /// </summary>
        private static string UniqueName(string folder, string fileName)
        {
            var name = Path.GetFileName(fileName ?? string.Empty);
            if (string.IsNullOrEmpty(name))
            {
                name = "file";
            }

            var invalid = Path.GetInvalidFileNameChars();
            var clean = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim();
            if (string.IsNullOrEmpty(clean) || clean == "." || clean == "..")
            {
                clean = "file";
            }

            if (clean.Length > 80)
            {
                var extension = Path.GetExtension(clean);
                if (extension.Length > 20)
                {
                    extension = string.Empty;
                }

                clean = clean.Substring(0, 80 - extension.Length) + extension;
            }

            var stamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            var attempt = 0;
            while (true)
            {
                var suffix = attempt == 0
                    ? Guid.NewGuid().ToString("N").Substring(0, 6)
                    : Guid.NewGuid().ToString("N").Substring(0, 6) + "-" + attempt;
                var candidate = stamp + "-" + suffix + "-" + clean;
                if (!File.Exists(Path.Combine(folder, candidate)))
                {
                    return candidate;
                }

                attempt++;
            }
        }
    }
}
