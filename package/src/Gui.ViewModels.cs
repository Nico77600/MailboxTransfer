// =============================================================================
//  Mailbox Transfer - console, view models of the browse component
// -----------------------------------------------------------------------------
//  Compiled by Initialize-MbtGuiTypes (MailboxTransfer.Gui.Core.ps1) with Add-Type, once per process. Not part of the engine unit: it has no
//  dependency on the engine and none on WPF (INotifyPropertyChanged only), so the engine stays free of GUI code.
//
//  FolderNode  a folder of the tree with a tri-state tick. Ticking a folder ticks the folders below it (the engine's IncludeSubfolders);
//              ticking some messages of a folder (Picks) makes it, and the folders above it, indeterminate. A folder whose children are not
//              read yet carries a placeholder; ticking it is remembered and given to the children when they are read.
//  MessageRow  a row of the list of a folder.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace MailboxTransferGui
{
    public class Bindable : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void Raise(string name) { var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(name)); }
    }

    public sealed class FolderNode : Bindable
    {
        public string Id;
        public string Glyph { get; set; }
        public string Sub { get; set; }
        /// <summary>The engine object (BrowseFolder) or the fake one.</summary>
        public object Tag;
        public bool IsRoot { get; set; }
        public string Kind { get; set; }
        /// <summary>Headers only (auxiliary archive): subjects, senders and text are not available before the export.</summary>
        public bool HeadersOnly;
        public FolderNode Parent;
        public ObservableCollection<FolderNode> Children { get; private set; }
        public readonly Dictionary<string, long> Picks = new Dictionary<string, long>();
        public readonly Dictionary<string, object> PickRows = new Dictionary<string, object>();
        /// <summary>Raised when the user opens a node whose children are not read yet.</summary>
        public static event Action<FolderNode> ExpandRequested;
        /// <summary>Raised when a tick changes (the selection tray and the footer follow).</summary>
        public static event Action SelectionChanged;

        string name; long count; long sizeBytes = -1; bool own; bool? check = false;
        bool expanded, selected, dim, loaded = true, loading;
        string status;

        public bool ChildrenNotRead { get { return !loaded; } }

        public string Name { get { return name; } set { name = value; Raise("Name"); } }
        public long Count { get { return count; } set { count = value; Raise("Count"); Raise("CountText"); } }
        public long SizeBytes { get { return sizeBytes; } set { sizeBytes = value; Raise("SizeBytes"); Raise("SizeText"); } }
        public string Status { get { return status; } set { status = value; Raise("Status"); } }
        public bool Dim { get { return dim; } set { dim = value; Raise("Dim"); } }
        public string CountText { get { return IsRoot || count <= 0 ? "" : count.ToString("N0", System.Globalization.CultureInfo.InvariantCulture); } }
        public string SizeText { get { return IsRoot || count <= 0 ? "" : (sizeBytes < 0 ? "size unknown" : FormatBytes(sizeBytes)); } }
        public bool HasPlaceholder { get { return Children.Count == 1 && Children[0].Id == "\u0000placeholder"; } }
        public bool ShowCheck { get { return !IsRoot; } }
        /// <summary>"Visible" / "Collapsed" as text: WPF converts it, and this file stays free of WPF types.</summary>
        public string CheckVisibility { get { return IsRoot || Id == "\u0000placeholder" ? "Collapsed" : "Visible"; } }
        public string SubVisibility { get { return string.IsNullOrEmpty(Sub) ? "Collapsed" : "Visible"; } }

        public bool IsExpanded
        {
            get { return expanded; }
            set
            {
                if (expanded == value) return;
                expanded = value; Raise("IsExpanded");
                if (value && !loaded && !loading) { loading = true; var h = ExpandRequested; if (h != null) h(this); }
            }
        }
        public bool IsSelected { get { return selected; } set { selected = value; Raise("IsSelected"); } }

        /// <summary>The folder (and the folders below it) ticked as a whole.</summary>
        public bool OwnAll { get { return own; } }

        /// <summary>True: the folder and everything below it; null: some of it; false: nothing. Two-way with the box.</summary>
        public bool? Check
        {
            get { return check; }
            set
            {
                if (IsRoot) return;
                SetAll(value == true);
                if (Parent != null) Parent.Refresh();
                Notify();
            }
        }

        public FolderNode(string id, string name, long count, long sizeBytes)
        {
            Id = id; this.name = name; this.count = count; this.sizeBytes = sizeBytes; Glyph = ""; Children = new ObservableCollection<FolderNode>();
        }

        public void MarkNotRead()
        {
            loaded = false; loading = false;
            Children.Clear();
            Children.Add(new FolderNode("\u0000placeholder", "Reading folders...", 0, -1) { Dim = true, Parent = this });
        }

        /// <summary>Gives the children read lazily; they inherit the tick of the folder.</summary>
        public void SetChildren(IEnumerable<FolderNode> kids)
        {
            Children.Clear();
            foreach (var k in kids) { k.Parent = this; Children.Add(k); if (own) k.SetAll(true); }
            loaded = true; loading = false;
            Raise("ChildrenNotRead");
            Refresh(); if (Parent != null) Parent.Refresh();
            Notify();
        }

        public void SetLoadFailed(string message)
        {
            Children.Clear();
            Children.Add(new FolderNode("\u0000placeholder", message, 0, -1) { Dim = true, Parent = this });
            loading = false;
        }

        void SetAll(bool on)
        {
            own = on;
            if (on) { Picks.Clear(); PickRows.Clear(); }
            foreach (var c in Children) if (c.Id != "\u0000placeholder") c.SetAll(on);
            if (!on) { Picks.Clear(); PickRows.Clear(); }
            Compute();
        }

        void Compute()
        {
            bool? v;
            var kids = Children.Where(c => c.Id != "\u0000placeholder").ToList();
            if (own && kids.All(c => c.check == true)) v = true;
            else if (own || Picks.Count > 0 || kids.Any(c => c.check != false)) v = null;
            else v = false;
            if (own && kids.Count > 0 && !kids.All(c => c.check == true)) v = null;
            if (check != v) { check = v; Raise("Check"); }
        }

        /// <summary>Recomputes this node and the nodes above it after a child changed.</summary>
        public void Refresh()
        {
            Compute();
            if (Parent != null) Parent.Refresh();
        }

        /// <summary>The user (un)ticks messages of this folder. A message ticked while the folder was ticked as a whole turns the folder into a list of picks.</summary>
        public void SetPicks(IEnumerable<KeyValuePair<string, object>> rows, Func<object, long> size, bool all, IEnumerable<string> allIds)
        {
            own = false; Picks.Clear(); PickRows.Clear();
            foreach (var kv in rows) { PickRows[kv.Key] = kv.Value; Picks[kv.Key] = size(kv.Value); }
            Refresh(); Notify();
        }

        public void TickWholeFolder(bool on)
        {
            SetAll(on);
            if (Parent != null) Parent.Refresh();
            Notify();
        }

        /// <summary>"Select all the messages of this folder": the messages of this folder only, not the folders below it.</summary>
        public void TickOwnOnly(bool on)
        {
            own = on; Picks.Clear(); PickRows.Clear();
            Compute();
            if (Parent != null) Parent.Refresh();
            Notify();
        }

        public static void Notify() { var h = SelectionChanged; if (h != null) h(); }

        public IEnumerable<FolderNode> Descendants()
        {
            yield return this;
            foreach (var c in Children) { if (c.Id == "\u0000placeholder") continue; foreach (var d in c.Descendants()) yield return d; }
        }

        public static string FormatBytes(long bytes)
        {
            if (bytes < 0) return "unknown";
            string[] units = { "B", "KiB", "MiB", "GiB", "TiB" };
            double v = bytes; int u = 0;
            while (v >= 1024 && u < units.Length - 1) { v /= 1024; u++; }
            if (u == 0) return v.ToString("0") + " B";
            string s = v >= 10 * 1024 * 1024 * 1024.0 || u < 3 ? v.ToString(u >= 3 && v >= 10 ? "0.0" : "0", System.Globalization.CultureInfo.InvariantCulture) : v.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
            if (u >= 3 && v < 10) s = v.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
            return s + " " + units[u];
        }
    }

    public sealed class MessageRow : Bindable
    {
        public string Id;
        public object Tag;
        public string From { get; set; }
        public string Address { get; set; }
        public string Subject { get; set; }
        public string SizeText { get; set; }
        public string DateText { get; set; }
        public long SizeBytes;
        public bool HasAttachment { get; set; }
        public string AttachmentGlyph { get { return HasAttachment ? "" : ""; } }
        public bool SubjectNa { get; set; }
        /// <summary>"Collapsed" for a list whose messages cannot be ticked one by one (a PST or .eml files to import: the folders are what is chosen).</summary>
        public string CheckVisibility { get; set; } = "Visible";
        bool isChecked;
        public Action<MessageRow> Toggled;
        public bool IsChecked
        {
            get { return isChecked; }
            set { if (isChecked == value) return; isChecked = value; Raise("IsChecked"); var h = Toggled; if (h != null) h(this); }
        }
        /// <summary>Sets the box without telling the folder (used when the list is built from the folder's picks).</summary>
        public void SetCheckedQuietly(bool value) { isChecked = value; Raise("IsChecked"); }
    }
}
