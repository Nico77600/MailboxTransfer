// =============================================================================
//  Mailbox Transfer - engine, part 17: choosing folders of a PST or of a folder of .eml files (-IncludeFolder / -ExcludeFolder)
// -----------------------------------------------------------------------------
//  FolderFilter   the folder paths the operator ticked or typed. A path is the slash separated path of the folder below the top of the
//                 source ("Inbox/Projects"; for a PST below "Top of Personal Folders", for .eml files below the chosen folder).
//
//  A pattern selects the folders whose path it matches (case does not matter) AND every folder below them. Wildcards: * is any text
//  (slashes included), ? is one character. "Inbox" is the Inbox and what is below it, "Inbox/*" what is below the Inbox but not the Inbox
//  itself, "*/2019" every folder called 2019 (and what is below it), "Archive*" every folder whose path starts with Archive.
//  A backtick makes the next character literal (`* is a star in the name of a folder). A folder is taken when it is selected by an include pattern (or there is no include pattern) and not selected by an exclude pattern:
//  an exclude always wins. To take a folder without the folders below it: include "Inbox", exclude "Inbox/*".
//  A folder that is not taken is left out; its subfolders are still looked at (they can be selected by their own pattern). The folders
//  above a taken folder are created in the target when they are not there (empty), as the tree needs them.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MailboxTransfer
{
    public sealed class FolderFilter
    {
        sealed class Pattern
        {
            public string Text;
            public Regex Rx;
            public bool Matched;
        }

        readonly List<Pattern> _include = new List<Pattern>(), _exclude = new List<Pattern>();

        public FolderFilter() { }

        public FolderFilter(IEnumerable<string> include, IEnumerable<string> exclude)
        {
            if (include != null) foreach (string p in include) AddInclude(p);
            if (exclude != null) foreach (string p in exclude) AddExclude(p);
        }

        /// <summary>A filter that takes everything (no pattern), or null when there is no pattern at all.</summary>
        public static FolderFilter Create(IEnumerable<string> include, IEnumerable<string> exclude)
        {
            var f = new FolderFilter(include, exclude);
            return f.IsActive ? f : null;
        }

        public bool IsActive { get { return _include.Count > 0 || _exclude.Count > 0; } }
        public IEnumerable<string> Includes { get { return _include.Select(p => p.Text); } }
        public IEnumerable<string> Excludes { get { return _exclude.Select(p => p.Text); } }

        public static string Normalize(string path)
        {
            return (path ?? "").Replace('\\', '/').Trim().Trim('/');
        }

        public void AddInclude(string pattern) { Add(_include, pattern); }
        public void AddExclude(string pattern) { Add(_exclude, pattern); }

        static void Add(List<Pattern> list, string pattern)
        {
            if (pattern == null) return;
            string t = Normalize(pattern);
            if (t.Length == 0 && pattern.Trim().Length == 0) return;
            var sb = new StringBuilder("^");
            for (int i = 0; i < t.Length; i++)
            {
                char c = t[i];
                if (c == '`' && i + 1 < t.Length) { sb.Append(Regex.Escape(t[++i].ToString())); }   // `* and `? are the characters themselves
                else if (c == '*') sb.Append(".*");
                else if (c == '?') sb.Append('.');
                else sb.Append(Regex.Escape(c.ToString()));
            }
            sb.Append('$');
            list.Add(new Pattern { Text = t, Rx = new Regex(sb.ToString(), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled) });
        }

        /// <summary>The path itself, or one of the folders above it, matches the pattern.</summary>
        static bool Selects(Pattern p, string path, bool mark)
        {
            string cur = path;
            while (true)
            {
                if (p.Rx.IsMatch(cur)) { if (mark) p.Matched = true; return true; }
                int i = cur.LastIndexOf('/');
                if (i < 0) return false;
                cur = cur.Substring(0, i);
            }
        }

        /// <summary>True when the folder at this path is taken. Remembers which patterns matched a folder (see <see cref="Unmatched"/>).</summary>
        public bool Allows(string path)
        {
            string p = Normalize(path);
            bool included = _include.Count == 0;
            foreach (Pattern x in _include) if (Selects(x, p, true)) included = true;
            bool excluded = false;
            foreach (Pattern x in _exclude) if (Selects(x, p, true)) excluded = true;
            return included && !excluded;
        }

        /// <summary>The folder is taken under any of these paths (the folder as the source names it, and as the file system names it).</summary>
        public bool AllowsAny(params string[] paths)
        {
            bool included = _include.Count == 0;
            foreach (string raw in paths)
            {
                string p = Normalize(raw);
                foreach (Pattern x in _include) if (Selects(x, p, true)) included = true;
            }
            bool excluded = false;
            foreach (string raw in paths)
            {
                string p = Normalize(raw);
                foreach (Pattern x in _exclude) if (Selects(x, p, true)) excluded = true;
            }
            return included && !excluded;
        }

        /// <summary>The include and exclude patterns that matched no folder so far (a typo in a path): "-IncludeFolder 'Inbx'".</summary>
        public List<string> Unmatched()
        {
            var l = new List<string>();
            foreach (Pattern p in _include) if (!p.Matched) l.Add("-IncludeFolder '" + p.Text + "'");
            foreach (Pattern p in _exclude) if (!p.Matched) l.Add("-ExcludeFolder '" + p.Text + "'");
            return l;
        }

        public override string ToString()
        {
            var parts = new List<string>();
            if (_include.Count > 0) parts.Add("include " + string.Join(", ", _include.Select(p => "'" + p.Text + "'")));
            if (_exclude.Count > 0) parts.Add("exclude " + string.Join(", ", _exclude.Select(p => "'" + p.Text + "'")));
            return parts.Count == 0 ? "every folder" : string.Join("; ", parts);
        }
    }
}
