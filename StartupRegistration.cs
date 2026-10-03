using System;
using System.IO;
using Microsoft.Win32;

namespace Imel
{
    /// <summary>
    /// Windows起動時の自動実行（HKCU の Run キー）への登録を管理します。
    /// </summary>
    internal static class StartupRegistration
    {
        private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string AppName = "Imel";

        internal static bool IsEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunKey, false);
                return key?.GetValue(AppName) != null;
            }
            catch
            {
                return false;
            }
        }

        // 失敗時は例外を投げ、呼び出し元で利用者に表示する。
        internal static void SetEnabled(bool enabled)
        {
            using var key = Registry.CurrentUser.CreateSubKey(RunKey, true);
            if (enabled)
            {
                string path = Environment.ProcessPath
                    ?? throw new InvalidOperationException("実行ファイルのパスを取得できません。");
                key.SetValue(AppName, $"\"{path}\"");
            }
            else
            {
                key.DeleteValue(AppName, false);
            }
        }

        // 登録済みの exe が移動・削除されている場合だけ、現在の exe で登録し直す。
        // 別の場所にある Imel を一時的に起動しただけで登録先を書き換えないよう、登録済みの exe が残っていれば変更しない。
        internal static void RefreshIfMoved()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RunKey, true);
                if (key?.GetValue(AppName) is not string command) return;

                string? current = Environment.ProcessPath;
                if (string.IsNullOrEmpty(current) || File.Exists(GetExecutablePath(command))) return;

                key.SetValue(AppName, $"\"{current}\"");
            }
            catch
            {
                // 自動実行の更新に失敗しても、アプリの起動は続ける。
            }
        }

        // Run キーの値（"C:\path\Imel.exe" または引用符なしのパス）から exe のパスを取り出す。
        internal static string GetExecutablePath(string command)
        {
            command = command.Trim();
            if (command.StartsWith('"'))
            {
                int end = command.IndexOf('"', 1);
                return end > 0 ? command.Substring(1, end - 1) : command.Substring(1);
            }
            return command;
        }
    }
}
