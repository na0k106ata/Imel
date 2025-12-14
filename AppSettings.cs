using System;
using System.IO;
using System.Text.Json;

namespace Imel
{
    /// <summary>
    /// アプリケーションの設定データを保持し、JSONファイルへの読み書きを行うクラス。
    /// </summary>
    public class AppSettings
    {
        // --- 位置・表示設定 ---
        public int OffsetX { get; set; } = 10;
        public int OffsetY { get; set; } = 10;
        public int Opacity { get; set; } = 67;
        public int UpdateInterval { get; set; } = 10; // デフォルト10ms

        // --- 動作設定 ---
        // マウスカーソルが非表示の時（動画視聴中など）に自動で隠すか
        public bool HideWhenCursorHidden { get; set; } = true;

        // --- 色設定 (RGB値を個別に保存) ---
        public byte TextR { get; set; } = 255;
        public byte TextG { get; set; } = 255;
        public byte TextB { get; set; } = 255;

        public byte BgR { get; set; } = 0;
        public byte BgG { get; set; } = 0;
        public byte BgB { get; set; } = 0;

        /// <summary>
        /// 設定ファイルのパスを取得します。(AppData/Roaming/Imel/settings.json)
        /// </summary>
        private static string GetConfigPath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "Imel");

            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }

            return Path.Combine(folder, "settings.json");
        }

        /// <summary>
        /// 設定をJSONファイルに保存します。
        /// </summary>
        public static void Save(AppSettings settings)
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string jsonString = JsonSerializer.Serialize(settings, options);
                File.WriteAllText(GetConfigPath(), jsonString);
            }
            catch { /* 保存失敗時は無視 */ }
        }

        /// <summary>
        /// 設定をJSONファイルから読み込みます。失敗時はデフォルト値を返します。
        /// </summary>
        public static AppSettings Load()
        {
            try
            {
                string path = GetConfigPath();
                if (File.Exists(path))
                {
                    string jsonString = File.ReadAllText(path);
                    var settings = JsonSerializer.Deserialize<AppSettings>(jsonString);
                    if (settings != null) return settings;
                }
            }
            catch { /* 読み込み失敗時は無視 */ }

            return new AppSettings();
        }
    }
}