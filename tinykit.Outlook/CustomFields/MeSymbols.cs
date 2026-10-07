using System;
using System.Collections.Generic;
using System.Globalization;

namespace tinykit.OutlookAddin.CustomFields
{
    /// <summary>A symbol the me column can show, with its code point and Unicode name for the chooser.</summary>
    internal sealed class MeSymbol
    {
        public string Text;
        public string Name;

        public MeSymbol(string text, string name)
        {
            Text = text;
            Name = name;
        }

        /// <summary>"U+25B6" (U+1F143 for a symbol beyond the BMP).</summary>
        public string CodePoint
        {
            get { return "U+" + char.ConvertToUtf32(Text, 0).ToString("X4", CultureInfo.InvariantCulture); }
        }

        public override string ToString()
        {
            return Text + "    " + CodePoint + "    " + Name;
        }
    }

    /// <summary>
    /// The me column's symbols: mail I sent, mail with me in To, mail with me in Cc (Mail settings meSent / meTo / meCc),
    /// and the choices the me Column Symbols window offers.
    /// </summary>
    internal static class MeSymbols
    {
        public const string DefaultSent = "▶";
        public const string DefaultTo = "●";
        public const string DefaultCc = "○";

        /// <summary>The current symbols (sent, to, cc); set by the controller from the mail settings.</summary>
        public static Func<Tuple<string, string, string>> Current = () => Tuple.Create(DefaultSent, DefaultTo, DefaultCc);

        public static readonly IList<MeSymbol> SentChoices = new[]
        {
            new MeSymbol("→", "RIGHTWARDS ARROW"),
            new MeSymbol("⇒", "RIGHTWARDS DOUBLE ARROW"),
            new MeSymbol("⇥", "RIGHTWARDS ARROW TO BAR"),
            new MeSymbol("⇨", "RIGHTWARDS WHITE ARROW"),
            new MeSymbol("▶", "BLACK RIGHT-POINTING TRIANGLE"),
            new MeSymbol("▷", "WHITE RIGHT-POINTING TRIANGLE"),
            new MeSymbol("⟶", "LONG RIGHTWARDS ARROW"),
            new MeSymbol("⟹", "LONG RIGHTWARDS DOUBLE ARROW"),
        };

        private static readonly MeSymbol[] Shapes =
        {
            new MeSymbol("■", "BLACK SQUARE"),
            new MeSymbol("□", "WHITE SQUARE"),
            new MeSymbol("▢", "WHITE SQUARE WITH ROUNDED CORNERS"),
            new MeSymbol("◆", "BLACK DIAMOND"),
            new MeSymbol("○", "WHITE CIRCLE"),
            new MeSymbol("●", "BLACK CIRCLE"),
            new MeSymbol("✓", "CHECK MARK"),
        };

        public static readonly IList<MeSymbol> ToChoices = With(Shapes,
            new MeSymbol("Ⓣ", "CIRCLED LATIN CAPITAL LETTER T"),
            new MeSymbol("\U0001F143", "SQUARED LATIN CAPITAL LETTER T"),
            new MeSymbol("\U0001F163", "NEGATIVE CIRCLED LATIN CAPITAL LETTER T"),
            new MeSymbol("\U0001F183", "NEGATIVE SQUARED LATIN CAPITAL LETTER T"));

        public static readonly IList<MeSymbol> CcChoices = With(Shapes,
            new MeSymbol("Ⓒ", "CIRCLED LATIN CAPITAL LETTER C"),
            new MeSymbol("\U0001F132", "SQUARED LATIN CAPITAL LETTER C"),
            new MeSymbol("\U0001F152", "NEGATIVE CIRCLED LATIN CAPITAL LETTER C"),
            new MeSymbol("\U0001F172", "NEGATIVE SQUARED LATIN CAPITAL LETTER C"));

        private static IList<MeSymbol> With(MeSymbol[] first, params MeSymbol[] more)
        {
            var list = new List<MeSymbol>(first);
            list.AddRange(more);
            return list;
        }
    }
}
