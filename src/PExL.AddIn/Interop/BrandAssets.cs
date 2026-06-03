using System.Drawing;
using System.Drawing.Drawing2D;

namespace PExL.AddIn.Interop
{
    /// <summary>
    /// Renders the PExL logo (the green→blue rounded tile with a white "|&gt;" pipe
    /// mark) at runtime via GDI+, so the ribbon shows the brand icon without shipping
    /// a bundled PNG. The geometry mirrors <c>brand/pexl-icon.svg</c> (a 256-unit
    /// design), scaled to whatever size the ribbon asks for.
    /// </summary>
    internal static class BrandAssets
    {
        private static Bitmap? _ribbonIcon;

        /// <summary>The icon for the "Open Editor" ribbon button (cached).</summary>
        public static Bitmap RibbonIcon() => _ribbonIcon ??= RenderIcon(32);

        /// <summary>Draw the app icon at the requested square size (the "light" variant).</summary>
        public static Bitmap RenderIcon(int size)
        {
            var bmp = new Bitmap(size, size);
            bmp.MakeTransparent();
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.Half;
                g.Clear(Color.Transparent);

                float s = size / 256f; // design is authored on a 256×256 grid

                // rounded tile with the brand gradient
                var tile = new RectangleF(8 * s, 8 * s, 240 * s, 240 * s);
                using (var path = RoundedRect(tile, 56 * s))
                using (var fill = new LinearGradientBrush(
                           tile, ColorTranslator.FromHtml("#4ec9b0"),
                           ColorTranslator.FromHtml("#569cd6"), 45f))
                {
                    g.FillPath(fill, path);
                }

                using (var white = new SolidBrush(Color.White))
                {
                    // the "|" bar
                    using (var bar = RoundedRect(new RectangleF(74 * s, 80 * s, 20 * s, 96 * s), 10 * s))
                        g.FillPath(white, bar);
                }

                // the ">" chevron
                using (var pen = new Pen(Color.White, 20 * s)
                {
                    StartCap = LineCap.Round,
                    EndCap = LineCap.Round,
                    LineJoin = LineJoin.Round
                })
                {
                    g.DrawLines(pen, new[]
                    {
                        new PointF(116 * s, 80 * s),
                        new PointF(168 * s, 128 * s),
                        new PointF(116 * s, 176 * s)
                    });
                }
            }
            return bmp;
        }

        private static GraphicsPath RoundedRect(RectangleF r, float radius)
        {
            float d = radius * 2f;
            var path = new GraphicsPath();
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
