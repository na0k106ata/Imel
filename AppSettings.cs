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

        /// <summary>
        /// 表示位置オフセット X (DIP)
        /// </summary>
        public int OffsetX { get; set; } = 10;

        /// <summary>
        /// 表示位置オフセット Y (DIP)
        /// </summary>
        public int OffsetY { get; set; } = 10;

        /// <summary>
        /// ウィンドウの不透明度 (0-100)
        /// </summary>
        public int Opacity { get; set; } = 67;

        /// <summary>
        /// インジケーターの位置更新間隔 (ミリ秒)
        /// </summary>
        public int UpdateInterval { get; set; } = 16;

        /// <summary>
        /// 表示サイズ倍率 (デフォルト: 1.0)
        /// </summary>
        public double Scale { get; set; } = 1.0;

        // --- 動作設定 ---

        /// <summary>
        /// OS側でマウスカーソルが非表示になった際にウィンドウを隠すかどうか
        /// </summary>
        public bool HideWhenCursorHidden { get; set; } = true;

        /// <summary>画面端でインジケーターの配置を反転するか（既定は無効）。</summary>
        public bool FlipAtScreenEdge { get; set; } = false;

        // --- 色設定 (RGB) ---

        /// <summary>
        /// 文字色の赤成分 (0-255)
        /// </summary>
        public byte TextR { get; set; } = 255;
        public byte TextG { get; set; } = 255;
        public byte TextB { get; set; } = 255;

        /// <summary>
        /// 背景色の赤成分 (0-255)
        /// </summary>
        public byte BgR { get; set; } = 0;
        public byte BgG { get; set; } = 0;
        public byte BgB { get; set; } = 0;

        private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

        private static string GetConfigPath() => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Imel", "settings.json");

        // 同じディレクトリ内で置換し、書き込み途中のファイルを本体として公開しない。
        public static bool Save(AppSettings settings, out string? error, string? path = null)
        {
            string? temporaryPath = null;
            try
            {
                path = Path.GetFullPath(path ?? GetConfigPath());
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                temporaryPath = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    JsonSerializer.Serialize(stream, settings, JsonOptions);
                    stream.Flush(flushToDisk: true);
                }

                if (File.Exists(path))
                {
                    // 破損した本体で正常なバックアップを上書きしない。
                    string? backup = TryRead(path) != null ? path + ".bak" : null;
                    File.Replace(temporaryPath, path, backup);
                }
                else
                {
                    File.Move(temporaryPath, path);
                }

                error = null;
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            finally
            {
                if (temporaryPath != null)
                {
                    try { File.Delete(temporaryPath); }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
            }
        }

        public static AppSettings Load(string? path = null)
        {
            path ??= GetConfigPath();
            return TryRead(path) ?? TryRead(path + ".bak") ?? new AppSettings();
        }

        private static AppSettings? TryRead(string path)
        {
            try
            {
                using var stream = File.OpenRead(path);
                var settings = JsonSerializer.Deserialize<AppSettings>(stream);
                if (settings == null) return null;
                settings.Scale = double.IsFinite(settings.Scale) ? Math.Clamp(settings.Scale, 0.5, 2.0) : 1.0;
                settings.Opacity = Math.Clamp(settings.Opacity, 0, 100);
                settings.UpdateInterval = Math.Clamp(settings.UpdateInterval, 2, 100);
                return settings;
            }
            catch (IOException) { return null; }
            catch (UnauthorizedAccessException) { return null; }
            catch (JsonException) { return null; }
        }
    }
}
