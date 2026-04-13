using System;
using System.IO;
using System.Text;

namespace WatchBox
{
    public static class ChangeLog
    {
        const string Header = "timestamp,action,item_id,name,details";

        public static void Append(string outputRoot, string action,
            string itemId, string name, string details = "")
        {
            if (string.IsNullOrEmpty(outputRoot)) return;
            var csvPath = Path.Combine(outputRoot, "log.csv");
            if (!File.Exists(csvPath))
                File.WriteAllText(csvPath, Header + Environment.NewLine,
                    new UTF8Encoding(true));

            string line = string.Join(",", new[] {
                DateTime.Now.ToString("yyyy-MM-dd\\THH:mm:ss"),
                action,
                CsvQuote(itemId),
                CsvQuote(name),
                CsvQuote(details)
            });
            using (var fs = new FileStream(csvPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
            using (var sw = new StreamWriter(fs, new UTF8Encoding(true)))
            {
                sw.WriteLine(line);
            }
        }

        // Quote a CSV field if it contains comma, quote, or newline (RFC 4180)
        static string CsvQuote(string value)
        {
            if (value == null) return "";
            if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0)
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            return value;
        }
    }
}
