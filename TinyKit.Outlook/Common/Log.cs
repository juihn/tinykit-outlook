using System;
using System.IO;
using TinyKit.OutlookAddin.Settings;

namespace TinyKit.OutlookAddin.Common
{
    /// <summary>Minimal append-only log in the settings folder (errors from background/event paths).</summary>
    internal static class Log
    {
        private const long MaxBytes = 1024 * 1024;

        public static void Error(string context, Exception ex)
        {
            Write("ERROR " + context + ": " + ex);
        }

        public static void Info(string message)
        {
            Write("INFO  " + message);
        }

        private static void Write(string line)
        {
            try
            {
                Directory.CreateDirectory(SettingsPaths.LocalFolder);
                var path = SettingsPaths.LogFile;
                var fi = new FileInfo(path);
                if (fi.Exists && fi.Length > MaxBytes)
                    File.Delete(path);
                File.AppendAllText(path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + line + Environment.NewLine);
            }
            catch
            {
                // logging must never throw
            }
        }
    }
}
