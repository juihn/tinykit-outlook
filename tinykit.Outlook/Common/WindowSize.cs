using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.Common
{
    /// <summary>
    /// A window's last size, kept in a small file of this PC ("width height" in pixels, e.g. Custom Filter Size.txt), so
    /// it opens as large as it was left.
    /// </summary>
    internal static class WindowSize
    {
        /// <summary>The saved size (no larger than the screen), or null when there is none.</summary>
        public static Size? Load(string path)
        {
            try
            {
                if (!File.Exists(path))
                    return null;
                var parts = File.ReadAllText(path).Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
                int w, h;
                if (parts.Length == 2 && int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out w)
                    && int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out h) && w > 0 && h > 0)
                {
                    var screen = Screen.PrimaryScreen.WorkingArea;
                    return new Size(Math.Min(w, screen.Width), Math.Min(h, screen.Height));
                }
            }
            catch (Exception ex)
            {
                Log.Error(Path.GetFileNameWithoutExtension(path), ex);
            }
            return null;
        }

        /// <summary>Keeps the window's size (its normal size when it is maximized).</summary>
        public static void Save(string path, Form form)
        {
            try
            {
                var size = form.WindowState == FormWindowState.Normal ? form.Size : form.RestoreBounds.Size;
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, size.Width.ToString(CultureInfo.InvariantCulture) + " " + size.Height.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                Log.Error(Path.GetFileNameWithoutExtension(path), ex);
            }
        }
    }
}
