using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace Imel
{
    internal static class ScreenPlacement
    {
        private const uint NearestMonitor = 2;
        private const uint NoZOrder = 0x0004, NoActivate = 0x0010;

        internal static bool PlaceIndicator(IntPtr window, int x, int y, double width, double height,
            double offsetX, double offsetY, bool flipAtScreenEdge)
        {
            if (window == IntPtr.Zero) return false;
            IntPtr monitor = MonitorFromPoint(new POINT { X = x, Y = y }, NearestMonitor);
            var info = new MONITORINFO { Size = Marshal.SizeOf<MONITORINFO>() };
            if (!GetMonitorInfo(monitor, ref info)) return false;
            var work = new PixelRect(info.Work.Left, info.Work.Top, info.Work.Right, info.Work.Bottom);
            // 移動で WM_DPICHANGED が同期通知された場合は、新しい DPI でもう一度配置する。
            for (int attempt = 0; attempt < 2; attempt++)
            {
                uint dpi = GetDpiForWindow(window);
                if (dpi == 0) return false;
                var rect = IndicatorPlacement.Calculate(x, y, width, height, offsetX, offsetY, dpi / 96.0, work, flipAtScreenEdge);
                if (GetWindowRect(window, out RECT current) && current.Left == rect.Left &&
                    current.Top == rect.Top && current.Right == rect.Right && current.Bottom == rect.Bottom)
                    return true;
                if (!SetWindowPos(window, IntPtr.Zero, rect.Left, rect.Top, rect.Width, rect.Height, NoZOrder | NoActivate))
                    return false;
                if (GetDpiForWindow(window) == dpi) return true;
            }
            return true;
        }

        internal static void FitSettingsWindow(Window window)
        {
            IntPtr handle = new WindowInteropHelper(window).Handle;
            if (handle == IntPtr.Zero) return;
            var info = new MONITORINFO { Size = Marshal.SizeOf<MONITORINFO>() };
            if (!GetMonitorInfo(MonitorFromWindow(handle, NearestMonitor), ref info) ||
                !GetWindowRect(handle, out RECT rect)) return;
            uint dpi = GetDpiForWindow(handle);
            if (dpi == 0) return;
            double scale = dpi / 96.0;
            int availableWidth = Math.Max(1, info.Work.Right - info.Work.Left);
            int availableHeight = Math.Max(1, info.Work.Bottom - info.Work.Top);
            window.MaxWidth = availableWidth / scale;
            window.MaxHeight = availableHeight / scale;
            int width = Math.Clamp(rect.Right - rect.Left, 1, availableWidth);
            int height = Math.Clamp(rect.Bottom - rect.Top, 1, availableHeight);
            int left = Math.Clamp(rect.Left, info.Work.Left, info.Work.Right - width);
            int top = Math.Clamp(rect.Top, info.Work.Top, info.Work.Bottom - height);
            SetWindowPos(handle, IntPtr.Zero, left, top, width, height, NoZOrder | NoActivate);
        }

        [StructLayout(LayoutKind.Sequential)] private struct POINT { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)] private struct RECT { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] private struct MONITORINFO
        {
            public int Size;
            public RECT Monitor, Work;
            public uint Flags;
        }
        [DllImport("user32.dll")] private static extern IntPtr MonitorFromPoint(POINT point, uint flags);
        [DllImport("user32.dll")] private static extern IntPtr MonitorFromWindow(IntPtr window, uint flags);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] private static extern bool GetMonitorInfo(IntPtr monitor, ref MONITORINFO info);
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out RECT rect);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr window, IntPtr after,
            int x, int y, int width, int height, uint flags);
    }
}
