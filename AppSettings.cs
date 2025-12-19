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

        /// <summary>
        /// ウィンドウの不透明度 (0-100)
        /// </summary>
        public int Opacity { get; set; } = 67;

        /// <summary>
        /// キャレット位置の監視間隔 (ミリ秒)
        /// </summary>
        public int UpdateInterval { get; set; } = 10;

        // 表示サイズ倍率 (デフォルト: 1.0)
        public double Scale { get; set; } = 1.0;

        // --- 動作設定 ---

        /// <summary>
        /// OS側でマウスカーソルが非表示になった際にウィンドウを隠すかどうか
        /// </summary>
        public bool HideWhenCursorHidden { get; set; } = true;

        // --- 色設定 (RGB) ---
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
            // Roamingフォルダを使用することで、ユーザーごとの設定として保存されます
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
            catch
            {
                // 保存失敗時は例外を無視します（ユーザー操作を妨げないため）
            }
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
            catch
            {
                // 読み込み失敗時（ファイル破損など）は無視してデフォルト設定を使用します
            }

            return new AppSettings();
        }
    }
}