using System;
using System.Collections.Generic;

namespace WatchBox
{
    public class ScanResult
    {
        public string ItemId;
        public string Name;
        public string SourcePath;
        public long Size;
        public DateTime ModifiedAt;
        public DateTime CreatedAt;
        // Mail-specific
        public string SenderEmail;
        public string SenderName;
        public string Subject;
        public DateTime ReceivedAt;
        public string BodyText;
        public List<string> AttachmentNames;
        // Folder where the item was materialized (set after Materialize)
        public string ItemFolder;
        // Paths set after materialization (mail-specific)
        public string BodyPath;
        public string MsgPath;
        public string AttachmentPaths;
        // Recipient addresses (semicolon-delimited SMTP addresses)
        public string ToRecipients;
        public string CcRecipients;
    }

    public class RunResult
    {
        public int Added;
        public int Modified;
        public int Removed;
    }

    public abstract class SourceScanner
    {
        public event Action<int, string> ProgressChanged;
        public volatile bool CancelRequested;

        protected void OnProgress(int count, string name)
        {
            if (ProgressChanged != null) ProgressChanged(count, name);
        }

        // Scan source and return new/changed items (not yet in knownIds, or changed)
        public abstract List<ScanResult> Scan(
            Dictionary<string, string> config, HashSet<string> knownIds);

        // Detect items that were removed from source
        public abstract List<string> DetectRemoved(
            Dictionary<string, string> config, HashSet<string> knownIds);
    }
}
