using System;

namespace Imel
{
    internal static class ImeStatusReader
    {
        internal delegate bool QueryControl(IntPtr imeWindow, int command, out IntPtr result);

        // null は取得不能。IME OFF（_A）とは区別して呼び出し元で非表示にする。
        internal static string? Read(IntPtr target, Func<IntPtr, IntPtr> getImeWindow, QueryControl query)
        {
            if (target == IntPtr.Zero) return null;
            IntPtr imeWindow = getImeWindow(target);
            if (imeWindow == IntPtr.Zero) return null;

            if (!query(imeWindow, 0x0005, out IntPtr open)) return null; // IMC_GETOPENSTATUS
            if (open == IntPtr.Zero) return "_A";
            if (!query(imeWindow, 0x0001, out IntPtr conversion)) return null; // IMC_GETCONVERSIONMODE

            int mode = conversion.ToInt32();
            if ((mode & 0x0001) != 0) // IME_CMODE_NATIVE
            {
                if ((mode & 0x0002) != 0) // IME_CMODE_KATAKANA
                    return (mode & 0x0008) != 0 ? "カ" : "_ｶ";
                return "あ";
            }

            return (mode & 0x0008) != 0 ? "Ａ" : "_A";
        }
    }
}
