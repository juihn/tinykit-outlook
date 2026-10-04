using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace tinykit.OutlookAddin.Common
{
    /// <summary>Office's built-in images (imageMso) as bitmaps for WinForms, with their transparency.</summary>
    internal static class OfficeImage
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct BITMAP
        {
            public int bmType, bmWidth, bmHeight, bmWidthBytes;
            public ushort bmPlanes, bmBitsPixel;
            public IntPtr bmBits;
        }

        [DllImport("gdi32.dll")]
        private static extern int GetObject(IntPtr h, int size, out BITMAP bitmap);

        // The gray of the ribbon's line icons; GetImageMso draws them in a light gray (#D4D4D4) that hardly shows on a
        // light window.
        private const int IconGray = 0x42;

        /// <summary>
        /// The image at <paramref name="size"/> pixels. GetImageMso returns a 32-bit DIB (rows top-down, alpha not
        /// premultiplied) whose alpha Image.FromHbitmap drops, so its pixels are copied directly; gray pixels are darkened
        /// to the ribbon's icon gray, colored ones (e.g. a task's red tick) are kept.
        /// </summary>
        public static Bitmap FromImageMso(Office.CommandBars bars, string imageMso, int size)
        {
            dynamic picture = bars.GetImageMso(imageMso, size, size);
            var handle = new IntPtr((int)picture.Handle);
            BITMAP info;
            if (GetObject(handle, Marshal.SizeOf(typeof(BITMAP)), out info) == 0 || info.bmBits == IntPtr.Zero || info.bmBitsPixel != 32)
                return Image.FromHbitmap(handle); // not a 32-bit DIB: no alpha to keep

            int width = info.bmWidth, height = info.bmHeight, stride = info.bmWidthBytes;
            var pixels = new byte[stride * height];
            Marshal.Copy(info.bmBits, pixels, 0, pixels.Length);
            for (int i = 0; i + 3 < pixels.Length; i += 4)
            {
                byte b = pixels[i], g = pixels[i + 1], r = pixels[i + 2];
                if (Math.Abs(r - g) < 8 && Math.Abs(g - b) < 8)
                    pixels[i] = pixels[i + 1] = pixels[i + 2] = IconGray;
            }
            var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            var data = bitmap.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
            try
            {
                for (int y = 0; y < height; y++)
                    Marshal.Copy(pixels, y * stride, data.Scan0 + y * data.Stride, width * 4);
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
            return bitmap;
        }
    }
}
