using System;
using System.Drawing;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace TinyKit.OutlookAddin.Formatting
{
    /// <summary>
    /// ViewFont.Color only takes the 16 OlColor values, so any RGB is mapped to the nearest one.
    /// </summary>
    internal static class OlColorMap
    {
        private sealed class Entry
        {
            public Outlook.OlColor Value;
            public string Name;
            public Color Rgb;

            public Entry(Outlook.OlColor value, string name, int rgb)
            {
                Value = value;
                Name = name;
                Rgb = Color.FromArgb((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
            }
        }

        private static readonly Entry[] Entries =
        {
            new Entry(Outlook.OlColor.olColorBlack, "Black", 0x000000),
            new Entry(Outlook.OlColor.olColorMaroon, "Maroon", 0x800000),
            new Entry(Outlook.OlColor.olColorGreen, "Green", 0x008000),
            new Entry(Outlook.OlColor.olColorOlive, "Olive", 0x808000),
            new Entry(Outlook.OlColor.olColorNavy, "Navy", 0x000080),
            new Entry(Outlook.OlColor.olColorPurple, "Purple", 0x800080),
            new Entry(Outlook.OlColor.olColorTeal, "Teal", 0x008080),
            new Entry(Outlook.OlColor.olColorGray, "Gray", 0x808080),
            new Entry(Outlook.OlColor.olColorSilver, "Silver", 0xC0C0C0),
            new Entry(Outlook.OlColor.olColorRed, "Red", 0xFF0000),
            new Entry(Outlook.OlColor.olColorLime, "Lime", 0x00FF00),
            new Entry(Outlook.OlColor.olColorYellow, "Yellow", 0xFFFF00),
            new Entry(Outlook.OlColor.olColorBlue, "Blue", 0x0000FF),
            new Entry(Outlook.OlColor.olColorFuchsia, "Fuchsia", 0xFF00FF),
            new Entry(Outlook.OlColor.olColorAqua, "Aqua", 0x00FFFF),
            new Entry(Outlook.OlColor.olColorWhite, "White", 0xFFFFFF),
        };

        public const string AutoName = "Auto";

        public static string ToName(Outlook.OlColor color)
        {
            var e = Entries.FirstOrDefault(x => x.Value == color);
            return e == null ? AutoName : e.Name;
        }

        /// <summary>Parses a palette name ("Red"), "Auto", or "#RRGGBB" (mapped to the nearest palette color).</summary>
        public static Outlook.OlColor Parse(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return Outlook.OlColor.olAutoColor;
            text = text.Trim();
            if (string.Equals(text, AutoName, StringComparison.OrdinalIgnoreCase))
                return Outlook.OlColor.olAutoColor;

            var named = Entries.FirstOrDefault(x => string.Equals(x.Name, text, StringComparison.OrdinalIgnoreCase));
            if (named != null)
                return named.Value;

            try
            {
                return Nearest(ColorTranslator.FromHtml(text));
            }
            catch (Exception)
            {
                return Outlook.OlColor.olAutoColor;
            }
        }

        public static Outlook.OlColor Nearest(Color color)
        {
            return Entries
                .OrderBy(e => Sq(e.Rgb.R - color.R) + Sq(e.Rgb.G - color.G) + Sq(e.Rgb.B - color.B))
                .First().Value;
        }

        /// <summary>The 16 palette colors in display order (without Auto).</summary>
        public static Outlook.OlColor[] Palette
        {
            get { return Entries.Select(e => e.Value).ToArray(); }
        }

        /// <summary>Screen color of a palette value; null for Auto.</summary>
        public static Color? ToColor(Outlook.OlColor color)
        {
            var e = Entries.FirstOrDefault(x => x.Value == color);
            return e == null ? (Color?)null : e.Rgb;
        }

        private static int Sq(int v)
        {
            return v * v;
        }
    }
}
