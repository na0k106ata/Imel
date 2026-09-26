using System;

namespace Imel
{
    internal readonly record struct PixelRect(int Left, int Top, int Right, int Bottom)
    {
        internal int Width => Right - Left;
        internal int Height => Bottom - Top;
    }

    internal static class IndicatorPlacement
    {
        // デスクトップ全体の物理座標で計算する。負座標のモニターにも対応。
        internal static PixelRect Calculate(int cursorX, int cursorY, double widthDip, double heightDip,
            double offsetXDip, double offsetYDip, double dpiScale, PixelRect work, bool flipAtScreenEdge = false)
        {
            int width = Math.Max(1, (int)Math.Ceiling(widthDip * dpiScale));
            int height = Math.Max(1, (int)Math.Ceiling(heightDip * dpiScale));
            double dx = offsetXDip * dpiScale, dy = offsetYDip * dpiScale;
            double x = cursorX + dx, y = cursorY + dy;
            if (!flipAtScreenEdge)
            {
                // OFF は通常のオフセットを維持し、画面端への吸着も行わない。
                int rawLeft = (int)Math.Round(Math.Clamp(x, int.MinValue, int.MaxValue - width));
                int rawTop = (int)Math.Round(Math.Clamp(y, int.MinValue, int.MaxValue - height));
                return new PixelRect(rawLeft, rawTop, rawLeft + width, rawTop + height);
            }
            width = Math.Min(width, Math.Max(1, work.Width));
            height = Math.Min(height, Math.Max(1, work.Height));
            if (x + width > work.Right) x = cursorX - Math.Abs(dx) - width;
            if (y + height > work.Bottom) y = cursorY - Math.Abs(dy) - height;
            int left = (int)Math.Round(Math.Clamp(x, work.Left, work.Right - width));
            int top = (int)Math.Round(Math.Clamp(y, work.Top, work.Bottom - height));
            return new PixelRect(left, top, left + width, top + height);
        }
    }
}
