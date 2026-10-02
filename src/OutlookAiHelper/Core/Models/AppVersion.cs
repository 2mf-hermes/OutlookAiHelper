using System;
using System.Globalization;

namespace OutlookAiHelper.Core.Models
{
    /// <summary>
    /// Minimal semantic version ("1.2.3", "v1.2.3-rc.1") used to compare the running
    /// build against a GitHub release tag. Deliberately lenient about missing fields
    /// ("1.0" == "1.0.0") and leading "v" so a hand-written tag still compares right.
    /// </summary>
    public sealed class AppVersion : IComparable<AppVersion>
    {
        private readonly int[] _parts;
        private readonly string _prerelease;

        private AppVersion(int[] parts, string prerelease)
        {
            _parts = parts;
            _prerelease = prerelease ?? string.Empty;
        }

        public int Major
        {
            get { return _parts[0]; }
        }

        public int Minor
        {
            get { return _parts[1]; }
        }

        public int Patch
        {
            get { return _parts[2]; }
        }

        public int Revision
        {
            get { return _parts[3]; }
        }

        public string Prerelease
        {
            get { return _prerelease; }
        }

        public bool IsPrerelease
        {
            get { return _prerelease.Length > 0; }
        }

        public string Normalized
        {
            get
            {
                var text = _parts[0] + "." + _parts[1] + "." + _parts[2];
                if (_parts[3] != 0)
                {
                    text = text + "." + _parts[3];
                }

                if (IsPrerelease)
                {
                    text = text + "-" + _prerelease;
                }

                return text;
            }
        }

        public static bool TryParse(string text, out AppVersion version)
        {
            version = null;
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            var raw = text.Trim();
            if (raw.Length == 0)
            {
                return false;
            }

            if (raw[0] == 'v' || raw[0] == 'V')
            {
                raw = raw.Substring(1);
            }

            var plus = raw.IndexOf('+');
            if (plus >= 0)
            {
                raw = raw.Substring(0, plus);
            }

            var prerelease = string.Empty;
            var dash = raw.IndexOf('-');
            if (dash >= 0)
            {
                prerelease = raw.Substring(dash + 1).Trim();
                raw = raw.Substring(0, dash);
            }

            if (raw.Length == 0)
            {
                return false;
            }

            var chunks = raw.Split('.');
            if (chunks.Length == 0 || chunks.Length > 4)
            {
                return false;
            }

            var parts = new int[4];
            for (var i = 0; i < chunks.Length; i++)
            {
                var chunk = chunks[i].Trim();
                if (chunk.Length == 0)
                {
                    return false;
                }

                int value;
                if (!int.TryParse(chunk, NumberStyles.None, CultureInfo.InvariantCulture, out value) || value < 0)
                {
                    return false;
                }

                parts[i] = value;
            }

            version = new AppVersion(parts, prerelease);
            return true;
        }

        public static AppVersion ParseOrNull(string text)
        {
            AppVersion parsed;
            return TryParse(text, out parsed) ? parsed : null;
        }

        public int CompareTo(AppVersion other)
        {
            if (other == null)
            {
                return 1;
            }

            for (var i = 0; i < 4; i++)
            {
                if (_parts[i] != other._parts[i])
                {
                    return _parts[i] < other._parts[i] ? -1 : 1;
                }
            }

            // Same numbers: a release outranks its own prereleases (semver rule 11.3).
            if (_prerelease.Length == 0 && other._prerelease.Length == 0)
            {
                return 0;
            }

            if (_prerelease.Length == 0)
            {
                return 1;
            }

            if (other._prerelease.Length == 0)
            {
                return -1;
            }

            return string.Compare(_prerelease, other._prerelease, StringComparison.OrdinalIgnoreCase);
        }

        public bool IsNewerThan(AppVersion other)
        {
            return CompareTo(other) > 0;
        }

        public override string ToString()
        {
            return Normalized;
        }
    }
}
