using System.Configuration;
using System.Data;
using System.Windows;
using Wpf.Ui.Appearance; // テーマ管理用

namespace Imel
{
    /// <summary>
    /// アプリケーションのエントリーポイント
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // 設定を読み込んでテーマを適用
            var settings = AppSettings.Load();
            ApplyTheme(settings.Theme);
        }

        /// <summary>
        /// 指定されたテーマをアプリケーション全体に適用します。
        /// </summary>
        public static void ApplyTheme(AppTheme theme)
        {
            switch (theme)
            {
                case AppTheme.Light:
                    ApplicationThemeManager.Apply(ApplicationTheme.Light);
                    break;
                case AppTheme.Dark:
                    ApplicationThemeManager.Apply(ApplicationTheme.Dark);
                    break;
                case AppTheme.Auto:
                default:
                    // システム設定に追従 (変更も検知)
                    ApplicationThemeManager.ApplySystemTheme();
                    break;
            }
        }
    }
}