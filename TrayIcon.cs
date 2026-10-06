using System;
using Drawing = System.Drawing;
using Forms = System.Windows.Forms;

namespace Imel
{
    /// <summary>
    /// タスクトレイのアイコン、右クリックメニュー、更新の通知を扱います。
    /// </summary>
    internal sealed class TrayIcon : IDisposable
    {
        private const string BaseText = "Imel (IME Indicator)";

        private readonly Forms.NotifyIcon _notifyIcon;
        private readonly Forms.ToolStripMenuItem _updateMenuItem;
        private readonly Forms.ToolStripSeparator _updateMenuSeparator;

        internal TrayIcon(Action openSettings, Action exit, Action openReleasePage)
        {
            _notifyIcon = new Forms.NotifyIcon
            {
                Icon = LoadAppIcon() ?? Drawing.SystemIcons.Application,
                Text = BaseText,
                Visible = true
            };
            _notifyIcon.BalloonTipClicked += (s, e) => openReleasePage();
            _notifyIcon.DoubleClick += (s, e) => openSettings();

            var contextMenu = new Forms.ContextMenuStrip();
            // 更新があるときだけ表示する。
            _updateMenuItem = new Forms.ToolStripMenuItem { Available = false };
            _updateMenuItem.Click += (s, e) => openReleasePage();
            _updateMenuSeparator = new Forms.ToolStripSeparator { Available = false };
            contextMenu.Items.Add(_updateMenuItem);
            contextMenu.Items.Add(_updateMenuSeparator);
            var settingsItem = new Forms.ToolStripMenuItem("設定...");
            settingsItem.Click += (s, e) => openSettings();
            var exitItem = new Forms.ToolStripMenuItem("終了");
            exitItem.Click += (s, e) => exit();

            contextMenu.Items.Add(settingsItem);
            contextMenu.Items.Add(new Forms.ToolStripSeparator());
            contextMenu.Items.Add(exitItem);
            _notifyIcon.ContextMenuStrip = contextMenu;
        }

        /// <summary>埋め込みリソースのアプリアイコンを読み込む。読み込めない場合は null。</summary>
        internal static Drawing.Icon? LoadAppIcon()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using var stream = assembly.GetManifestResourceStream("Imel.Imel.ico");
                return stream != null ? new Drawing.Icon(stream) : null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Imel: アプリのアイコンを読み込めません: {ex}");
                return null;
            }
        }

        /// <summary>更新があればメニューとツールチップに反映し、なければ元に戻す。</summary>
        internal void ShowUpdate(UpdateInfo? update)
        {
            _updateMenuItem.Available = update != null;
            _updateMenuSeparator.Available = update != null;
            _notifyIcon.Text = update != null ? BaseText + " - 更新があります" : BaseText;
            if (update != null) _updateMenuItem.Text = $"v{update.Version}をダウンロード...";
        }

        internal void ShowUpdateBalloon(UpdateInfo update, Version currentVersion)
        {
            _notifyIcon.ShowBalloonTip(10000, "Imelの新しいバージョンがあります",
                $"v{update.Version}が公開されました（現在 v{currentVersion}）。クリックするとダウンロードページを開きます。",
                Forms.ToolTipIcon.Info);
        }

        public void Dispose()
        {
            _notifyIcon.ContextMenuStrip?.Dispose();
            _notifyIcon.Dispose();
        }
    }
}
