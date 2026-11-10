// Test and lab helper of the PST repair: after a repair, every item the ordinary reader could read from the damaged source must be in the
// target, in the same folder and identical (PstFxCompare). Compiled in the same unit as the engine (see Import-StoreEngine.ps1).
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MailboxTransfer.TestSupport
{
    public sealed class RepairVerification
    {
        public int Readable, Present, Identical, TargetMessages, SourceMessagesListed;
        public readonly List<string> Problems = new List<string>();
        public bool Ok { get { return Problems.Count == 0; } }
    }

    public static class PstRepairVerifier
    {
        static void Touch(FxProperty p)
        {
            if (!p.IsLazy) return;
            byte[] buf = new byte[65536];
            using (Stream s = p.OpenValue()) { while (s.Read(buf, 0, buf.Length) > 0) { } }
        }

        static void Touch(FxMessage m)
        {
            foreach (FxProperty p in m.Properties) Touch(p);
            foreach (FxRecipient r in m.Recipients) foreach (FxProperty p in r.Properties) Touch(p);
            foreach (FxAttachment a in m.Attachments)
            {
                foreach (FxProperty p in a.Properties) Touch(p);
                if (a.Embedded != null) Touch(a.Embedded);
            }
        }

        static void Paths(PstReader r, Dictionary<uint, string> into)
        {
            foreach (PstReaderFolder f in r.AllFolders())
            {
                foreach (bool assoc in new[] { false, true })
                    foreach (uint nid in r.MessageNids(f, assoc)) into[nid] = f.Path;
            }
        }

        /// <summary>The items readable from the source by the ordinary reader (through the contents tables), checked in the target.</summary>
        public static RepairVerification Verify(string source, string target, PstRepairReport report)
        {
            var v = new RepairVerification();
            var map = new Dictionary<uint, PstRepairItem>();
            foreach (PstRepairItem it in report.Items) map[it.SourceNid] = it;
            var srcPaths = new Dictionary<uint, string>();
            PstReader sr = null;
            try
            {
                sr = PstReader.Open(source);
                try { Paths(sr, srcPaths); } catch (Exception) { }
            }
            catch (Exception) { return v; }
            var dstPaths = new Dictionary<uint, string>();
            using (PstReader tr = PstReader.Open(target))
            {
                Paths(tr, dstPaths);
                v.TargetMessages = dstPaths.Count;
                v.SourceMessagesListed = srcPaths.Count;
                foreach (KeyValuePair<uint, string> kv in srcPaths)
                {
                    FxMessage m;
                    try { m = sr.ReadMessage(kv.Key); Touch(m); }
                    catch (Exception) { continue; }                       // not readable before: not required
                    v.Readable++;
                    PstRepairItem it;
                    if (!map.TryGetValue(kv.Key, out it) || it.TargetNid == 0) { v.Problems.Add("0x" + kv.Key.ToString("X") + " [" + kv.Value + "] was readable and is not in the target"); continue; }
                    string tp;
                    if (!dstPaths.TryGetValue(it.TargetNid, out tp)) { v.Problems.Add("0x" + kv.Key.ToString("X") + " is not listed in any folder of the target"); continue; }
                    v.Present++;
                    if (!string.Equals(tp, kv.Value, StringComparison.Ordinal)) { v.Problems.Add("0x" + kv.Key.ToString("X") + " moved from '" + kv.Value + "' to '" + tp + "'"); continue; }
                    var diffs = new List<string>();
                    FxMessage c = tr.ReadMessage(it.TargetNid);
                    if (!PstFxCompare.Compare(m, c, diffs)) { v.Problems.Add("0x" + kv.Key.ToString("X") + " [" + kv.Value + "] differs: " + string.Join(" | ", diffs.Take(3))); continue; }
                    v.Identical++;
                }
            }
            sr.Dispose();
            return v;
        }
    }
}

namespace MailboxTransfer.TestSupport
{
    /// <summary>Progress that records the reports on the calling thread (System.Progress posts to the thread pool, where a script block has no runspace).</summary>
    public sealed class SyncProgress<T> : IProgress<T>
    {
        public readonly List<T> Reports = new List<T>();
        public void Report(T value) { lock (Reports) Reports.Add(value); }
    }
}
