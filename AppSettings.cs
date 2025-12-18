using System;
using System.IO;
using System.Text.Json;

namespace Imel
{
    /// <summary>
    /// アプリケーションのテーマ設定
    /// </summary>
    public enum AppTheme
    {
        Light,
        Dark,
        Auto
    }

    /// <summary>
    /// アプリケーション設定のデータモデルと保存・読み込みロジックを提供します。
    /// </summary>
    public class AppSettings
    {
        // 表示位置オフセット X (ピクセル)
        public int OffsetX { get; set; } = 10;
        // 表示位置オフセット Y (ピクセル)
        public int OffsetY { get; set; } = 10;

        // 背景の不透明度 (0-100%)
        public int Opacity { get; set; } = 67;

        // 更新間隔 (ミリ秒)
        public int UpdateInterval { get; set; } = 10;

        // OSのマウスカーソル非表示時に連動して隠すか
        public bool HideWhenCursorHidden { get; set; } = true;

        // 表示倍率 (0.5 - 2.0)
        public double Scale { get; set; } = 1.0;

        // テキスト色 (RGB)
        public byte TextR { get; set; } = 255;
        public byte TextG { get; set; } = 255;
        public byte TextB { get; set; } = 255;

        // 背景色 (RGB)
        public byte BgR { get; set; } = 0;
        public byte BgG { get; set; } = 0;
        public byte BgB { get; set; } = 0;

        // アプリケーションのテーマ (Light, Dark, Auto)
        public AppTheme Theme { get; set; } = AppTheme.Auto;

        // 設定ファイルの保存パス (AppDataフォルダ内)
        private static string SettingsPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Imel",
            "settings.json");

        /// <summary>
        /// 設定をJSONファイルとして保存します。
        /// </summary>
        public static void Save(AppSettings settings)
        {
            try
            {
                string dir = Path.GetDirectoryName(SettingsPath) ?? "";
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsPath, json);
            }
            catch { }
        }

        /// <summary>
        /// JSONファイルから設定を読み込みます。失敗した場合はデフォルト値を返します。
        /// </summary>
        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    string json = File.ReadAllText(SettingsPath);
                    var obj = JsonSerializer.Deserialize<AppSettings>(json);
                    return obj ?? new AppSettings();
                }
            }
            catch { }
            return new AppSettings();
        }
    }
}