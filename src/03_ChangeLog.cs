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
                CsvSafe(itemId),
                CsvSafe(name),
                CsvSafe(details)
            });
            File.AppendAllText(csvPath, line + Environment.NewLine, new UTF8Encoding(true));
        }

        static string CsvSafe(string value)
        {
            return (value ?? "").Replace(",", " ").Replace("\r", " ").Replace("\n", " ");
        }
    }
}
