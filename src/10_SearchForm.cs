using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;

namespace WatchBox
{
    public class SearchWindow : WebViewHost
    {
        public SearchWindow()
            : base("Viewer", "viewer.html", 1040, 680)
        {
            MinWidth = 700; MinHeight = 400;
        }

        protected override void HandleMessage(string action, string json)
        {
            switch (action)
            {
                case "getProfiles": SendProfiles(); break;
                case "getManifest": SendManifest(json); break;
                case "getMailBody": SendMailBody(json); break;
                case "getAttachments": SendAttachments(json); break;
                case "getFilePreview": SendFilePreview(json); break;
                case "getLog": SendLog(json); break;
                case "getConfigValue": SendConfigValue(json); break;
                case "setConfigValue": SetConfigValue(json); break;
                case "openFile": OpenFile(json); break;
                case "openDirectory": OpenDirectory(json); break;
            }
        }

        // --- Log data (log.csv) ---

        void SendLog(string json)
        {
            string outputRoot = ExtractJsonString(json, "outputRoot");
            var sb = new StringBuilder();
            sb.Append("{\"rows\":[");

            if (!string.IsNullOrEmpty(outputRoot))
            {
                string logPath = Path.Combine(outputRoot, "log.csv");
                if (File.Exists(logPath))
                {
                    var logLines = new List<string>();
                    using (var fs = new FileStream(logPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var sr = new StreamReader(fs, Encoding.UTF8))
                    {
                        string l;
                        while ((l = sr.ReadLine()) != null) logLines.Add(l);
                    }
                    bool first = true;
                    foreach (var line in logLines)
                    {
                        if (string.IsNullOrEmpty(line)) continue;
                        var cols = ManifestIO.CsvSplit(line);
                        if (cols.Length < 4 || cols[0] == "timestamp") continue;
                        if (!first) sb.Append(",");
                        first = false;
                        sb.AppendFormat("{{\"ts\":\"{0}\",\"action\":\"{1}\",\"id\":\"{2}\",\"name\":\"{3}\"}}",
                            JsonEsc(cols[0]), JsonEsc(cols[1]), JsonEsc(cols[2]),
                            cols.Length > 3 ? JsonEsc(cols[3]) : "");
                    }
                }
            }
            sb.Append("]}");
            SendMessage("logLoaded", sb.ToString());
        }

        void SendConfigValue(string json)
        {
            string key = ExtractJsonString(json, "key");
            string value = Config.Get(key, "");
            SendMessage("configValue",
                "{\"key\":\"" + JsonEsc(key) + "\",\"value\":\"" + JsonEsc(value) + "\"}");
        }

        void SetConfigValue(string json)
        {
            string key = ExtractJsonString(json, "key");
            string value = ExtractJsonString(json, "value");
            Config.Set(key, value);
            Config.Save();
        }

        // --- Send profile list ---

        void SendProfiles()
        {
            var sb = new StringBuilder();
            sb.Append("{\"profiles\":[");
            for (int i = 0; i < Config.ProfileCount; i++)
            {
                if (i > 0) sb.Append(",");
                sb.AppendFormat("{{\"index\":{0},\"name\":\"{1}\",\"type\":\"{2}\"}}",
                    i,
                    JsonEsc(Config.PGet(i, "name", "Profile " + (i + 1))),
                    JsonEsc(Config.PGet(i, "type", "mail")));
            }
            sb.Append("]}");
            SendMessage("profilesLoaded", sb.ToString());
        }

        // --- Send manifest data as JSON ---

        void SendManifest(string json)
        {
            int profileIndex = 0;
            var m = System.Text.RegularExpressions.Regex.Match(json, "\"profileIndex\"\\s*:\\s*(\\d+)");
            if (m.Success) profileIndex = int.Parse(m.Groups[1].Value);

            string type = Config.PGet(profileIndex, "type", "mail");
            string outputRoot = Config.PGet(profileIndex, "output_root");
            if (string.IsNullOrEmpty(outputRoot))
            {
                SendMessage("manifestLoaded",
                    "{\"type\":\"" + type + "\",\"outputRoot\":\"\",\"rows\":[]}");
                return;
            }

            string csvPath = ManifestIO.ResolvePath(outputRoot);
            bool isMail = type == "mail";

            var sb = new StringBuilder();
            sb.AppendFormat("{{\"type\":\"{0}\",\"outputRoot\":\"{1}\",\"rows\":[",
                type, JsonEsc(outputRoot));

            if (File.Exists(csvPath))
            {
                string[] lines;
                using (var fs = new FileStream(csvPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var sr = new StreamReader(fs, Encoding.UTF8))
                {
                    var lineList = new System.Collections.Generic.List<string>();
                    string l;
                    while ((l = sr.ReadLine()) != null) lineList.Add(l);
                    lines = lineList.ToArray();
                }
                bool first = true;
                foreach (var line in lines)
                {
                    if (string.IsNullOrEmpty(line)) continue;
                    var cols = ManifestIO.CsvSplit(line);
                    if (cols.Length < 2) continue;
                    if (cols[0] == "entry_id" || cols[0] == "item_id") continue;

                    if (!first) sb.Append(",");
                    first = false;

                    if (isMail)
                        AppendMailRow(sb, cols);
                    else
                        AppendFolderRow(sb, cols);
                }
            }

            sb.Append("]}");
            SendMessage("manifestLoaded", sb.ToString());
        }

        void AppendMailRow(StringBuilder sb, string[] cols)
        {
            // entry_id,sender_email,sender_name,subject,received_at,
            // folder_path,body_path,msg_path,attachment_paths,mail_folder,body_text,
            // to_recipients,cc_recipients
            sb.Append("{");
            sb.AppendFormat("\"entry_id\":\"{0}\"", JsonEsc(Col(cols, 0)));
            sb.AppendFormat(",\"sender_email\":\"{0}\"", JsonEsc(Col(cols, 1)));
            sb.AppendFormat(",\"sender_name\":\"{0}\"", JsonEsc(Col(cols, 2)));
            sb.AppendFormat(",\"subject\":\"{0}\"", JsonEsc(Col(cols, 3)));
            sb.AppendFormat(",\"received_at\":\"{0}\"", JsonEsc(Col(cols, 4)));
            sb.AppendFormat(",\"folder_path\":\"{0}\"", JsonEsc(Col(cols, 5)));
            sb.AppendFormat(",\"body_path\":\"{0}\"", JsonEsc(Col(cols, 6)));
            sb.AppendFormat(",\"msg_path\":\"{0}\"", JsonEsc(Col(cols, 7)));
            sb.AppendFormat(",\"attachment_paths\":\"{0}\"", JsonEsc(Col(cols, 8)));
            sb.AppendFormat(",\"mail_folder\":\"{0}\"", JsonEsc(Col(cols, 9)));
            sb.AppendFormat(",\"body_text\":\"{0}\"", JsonEsc(Col(cols, 10)));
            sb.AppendFormat(",\"to_recipients\":\"{0}\"", JsonEsc(Col(cols, 11)));
            sb.AppendFormat(",\"cc_recipients\":\"{0}\"", JsonEsc(Col(cols, 12)));
            sb.Append("}");
        }

        void AppendFolderRow(StringBuilder sb, string[] cols)
        {
            // item_id,file_name,file_path,folder_path,relative_path,file_size,modified_at
            sb.Append("{");
            sb.AppendFormat("\"item_id\":\"{0}\"", JsonEsc(Col(cols, 0)));
            sb.AppendFormat(",\"file_name\":\"{0}\"", JsonEsc(Col(cols, 1)));
            sb.AppendFormat(",\"file_path\":\"{0}\"", JsonEsc(Col(cols, 2)));
            sb.AppendFormat(",\"folder_path\":\"{0}\"", JsonEsc(Col(cols, 3)));
            sb.AppendFormat(",\"relative_path\":\"{0}\"", JsonEsc(Col(cols, 4)));
            sb.AppendFormat(",\"file_size\":\"{0}\"", JsonEsc(Col(cols, 5)));
            sb.AppendFormat(",\"modified_at\":\"{0}\"", JsonEsc(Col(cols, 6)));
            sb.Append("}");
        }

        // --- Mail body ---

        void SendMailBody(string json)
        {
            string bodyPath = ExtractJsonString(json, "bodyPath");
            string bodyText = ExtractJsonString(json, "bodyText");

            string body = bodyText;
            try
            {
                if (!string.IsNullOrEmpty(bodyPath) && File.Exists(bodyPath))
                    body = File.ReadAllText(bodyPath, Encoding.UTF8);
            }
            catch { }

            SendMessage("mailBodyLoaded",
                "{\"body\":\"" + JsonEsc(body ?? "") + "\"}");
        }

        // --- Attachments ---

        void SendAttachments(string json)
        {
            string mailFolder = ExtractJsonString(json, "mailFolder");
            var sb = new StringBuilder();
            sb.Append("{\"files\":[");

            if (!string.IsNullOrEmpty(mailFolder) && Directory.Exists(mailFolder))
            {
                try
                {
                    var files = Directory.GetFiles(mailFolder);
                    bool first = true;
                    foreach (var f in files)
                    {
                        string fn = Path.GetFileName(f);
                        if (fn == "meta.json" || fn == "body.txt" || fn == "mail.msg")
                            continue;
                        if (!first) sb.Append(",");
                        first = false;
                        sb.AppendFormat("{{\"name\":\"{0}\",\"path\":\"{1}\"}}",
                            JsonEsc(fn), JsonEsc(f));
                    }
                }
                catch { }
            }

            sb.Append("]}");
            SendMessage("attachmentsLoaded", sb.ToString());
        }

        // --- File preview ---

        void SendFilePreview(string json)
        {
            string filePath = ExtractJsonString(json, "filePath");
            string fileName = ExtractJsonString(json, "fileName");
            string previewType = ExtractJsonString(json, "previewType");

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                SendMessage("filePreview",
                    "{\"type\":\"error\",\"content\":\"File not found\"}");
                return;
            }

            try
            {
                switch (previewType)
                {
                    case "pdf":
                        SendVirtualPathPreview(filePath, previewType);
                        break;
                    case "image":
                        SendImagePreview(filePath, fileName);
                        break;
                    case "excel":
                    case "docx":
                    case "pptx":
                        SendBinaryPreview(filePath, previewType, fileName);
                        break;
                    case "text":
                    case "html":
                    case "markdown":
                        SendTextPreview(filePath, previewType);
                        break;
                    default:
                        SendMessage("filePreview",
                            "{\"type\":\"none\"}");
                        break;
                }
            }
            catch
            {
                SendMessage("filePreview",
                    "{\"type\":\"error\",\"content\":\"Preview failed\"}");
            }
        }

        void SendVirtualPathPreview(string filePath, string type)
        {
            // Map the file's directory so WebView2 can access it
            string dir = Path.GetDirectoryName(filePath);
            string fn = Path.GetFileName(filePath);
            string hostName = "preview.local";

            Dispatcher.Invoke(new Action(() =>
            {
                try
                {
                    _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        hostName, dir,
                        Microsoft.Web.WebView2.Core
                            .CoreWebView2HostResourceAccessKind.Allow);
                }
                catch { }
            }));

            string virtualPath = "https://" + hostName + "/" +
                Uri.EscapeDataString(fn);
            SendMessage("filePreview",
                "{\"type\":\"" + JsonEsc(type) +
                "\",\"virtualPath\":\"" + JsonEsc(virtualPath) + "\"}");
        }

        void SendImagePreview(string filePath, string fileName)
        {
            byte[] bytes = File.ReadAllBytes(filePath);
            string base64 = Convert.ToBase64String(bytes);
            string ext = Path.GetExtension(fileName).TrimStart('.').ToLower();
            // Map extension to MIME subtype
            string mime = ext;
            if (ext == "jpg") mime = "jpeg";
            else if (ext == "svg") mime = "svg+xml";
            else if (ext == "ico") mime = "x-icon";
            SendMessage("filePreview",
                "{\"type\":\"image\",\"ext\":\"" + JsonEsc(mime) +
                "\",\"content\":\"" + base64 + "\"}");
        }

        void SendBinaryPreview(string filePath, string type, string fileName)
        {
            byte[] bytes = File.ReadAllBytes(filePath);
            string base64 = Convert.ToBase64String(bytes);
            string ext = Path.GetExtension(fileName).TrimStart('.').ToLower();
            SendMessage("filePreview",
                "{\"type\":\"" + JsonEsc(type) +
                "\",\"ext\":\"" + JsonEsc(ext) +
                "\",\"content\":\"" + base64 + "\"}");
        }

        void SendTextPreview(string filePath, string type = "text")
        {
            string content = "";
            try
            {
                content = File.ReadAllText(filePath, Encoding.UTF8);
                if (content.Length > 102400)
                    content = content.Substring(0, 102400) + "\n\n... (truncated)";
            }
            catch { content = "(unable to read file)"; }

            SendMessage("filePreview",
                "{\"type\":\"" + JsonEsc(type) +
                "\",\"content\":\"" + JsonEsc(content) + "\"}");
        }

        // --- Open file ---

        void OpenFile(string json)
        {
            string path = ExtractJsonString(json, "path");
            if (!string.IsNullOrEmpty(path))
            {
                try { System.Diagnostics.Process.Start(path); }
                catch { }
            }
        }

        void OpenDirectory(string json)
        {
            string path = ExtractJsonString(json, "path");
            if (!string.IsNullOrEmpty(path))
            {
                string dir = Directory.Exists(path) ? path : Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                {
                    try { System.Diagnostics.Process.Start("explorer.exe", dir); }
                    catch { }
                }
            }
        }

        // --- Helpers ---

        static string Col(string[] cols, int idx)
        {
            return idx < cols.Length ? cols[idx] : "";
        }
    }
}
