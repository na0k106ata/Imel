using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Interop;
using System.Windows.Threading;
using System.Diagnostics;
using System.Threading.Tasks;
using Forms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace Imel
{
    /// <summary>
    /// アプリケーションのメインロジック。
    /// 透明なオーバーレイウィンドウを制御し、アクティブウィンドウのIME状態監視、
    /// マウスカーソル位置追従、タスクトレイアイコンの管理を行います。
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Fields

        private DispatcherTimer _timer;
        private Forms.NotifyIcon _notifyIcon = null!;
        private ImageSource? _cachedIcon = null;

        // 処理頻度の制御用（負荷軽減）
        private DateTime _lastImeCheckTime = DateTime.MinValue;
        private const double ImeCheckInterval = 100.0; // IME状態確認は100ms間隔

        private DateTime _lastMemoryCleanupTime = DateTime.MinValue;
        private const double MemoryCleanupInterval = 30000.0; // メモリ解放は30秒間隔

        // パフォーマンス最適化用キャッシュ
        private double _dpiX = 1.0;
        private double _dpiY = 1.0;
        private IntPtr _thisProcessHandle;

        #endregion

        #region Properties (Settings)

        // 位置調整用オフセット
        public int SettingOffsetX { get; set; } = 10;
        public int SettingOffsetY { get; set; } = 10;

        // カーソル非表示時の連動設定
        public bool SettingHideWhenCursorHidden { get; set; } = true;

        // 更新間隔（描画頻度）
        private int _settingUpdateInterval = 10;
        public int SettingUpdateInterval
        {
            get => _settingUpdateInterval;
            set
            {
                // 2ms~100msの範囲に制限
                _settingUpdateInterval = Math.Clamp(value, 2, 100);
                if (_timer != null)
                {
                    _timer.Interval = TimeSpan.FromMilliseconds(_settingUpdateInterval);
                }
            }
        }

        private Color _settingTextColor = Colors.White;
        public Color SettingTextColor
        {
            get => _settingTextColor;
            set
            {
                _settingTextColor = value;
                if (ImeStatusText != null)
                {
                    ImeStatusText.Foreground = new SolidColorBrush(value);
                }
            }
        }

        private Color _settingBackgroundColor = Colors.Black;
        public Color SettingBackgroundColor
        {
            get => _settingBackgroundColor;
            set
            {
                _settingBackgroundColor = value;
                UpdateBackgroundBrush();
            }
        }

        private int _settingOpacity = 67;
        public int SettingOpacity
        {
            get => _settingOpacity;
            set
            {
                _settingOpacity = Math.Clamp(value, 0, 100);
                UpdateBackgroundBrush();
            }
        }

        #endregion

        #region Initialization & Cleanup

        public MainWindow()
        {
            // 【メモリ対策】GPU描画を無効化しソフトウェア描画のみにする
            RenderOptions.ProcessRenderMode = RenderMode.SoftwareOnly;

            InitializeComponent();

            // プロセスハンドルをキャッシュ（メモリ解放用）
            using (var proc = Process.GetCurrentProcess())
            {
                _thisProcessHandle = proc.Handle;
            }

            LoadSettings();

            UpdateBackgroundBrush();
            if (ImeStatusText != null)
                ImeStatusText.Foreground = new SolidColorBrush(SettingTextColor);

            InitializeNotifyIcon();

            // アイコン画像をキャッシュしてGDIリソース漏れを防ぐ
            CacheAppIcon();

            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromMilliseconds(SettingUpdateInterval);
            _timer.Tick += Timer_Tick;

            this.Loaded += MainWindow_Loaded;
            this.Closing += MainWindow_Closing;
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // 初期位置を画面外へ（チラつき防止）
            this.Left = -100;
            this.Top = -100;

            // DPI情報の初期取得
            var source = PresentationSource.FromVisual(this);
            if (source != null && source.CompositionTarget != null)
            {
                _dpiX = source.CompositionTarget.TransformToDevice.M11;
                _dpiY = source.CompositionTarget.TransformToDevice.M22;
            }

            _timer.Start();

            // 【メモリ対策】起動直後の初期化メモリゴミを掃除するため、少し遅延させて解放実行
            await Task.Delay(2000);
            MinimizeFootprint();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveSettings();
            _notifyIcon.Dispose();
        }

        // DPI変更時の対応（モニター移動など）
        protected override void OnDpiChanged(DpiScale oldDpi, DpiScale newDpi)
        {
            base.OnDpiChanged(oldDpi, newDpi);
            _dpiX = newDpi.DpiScaleX;
            _dpiY = newDpi.DpiScaleY;
        }

        private void CacheAppIcon()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("Imel.Imel.ico"))
                {
                    if (stream != null)
                    {
                        using (var icon = new Drawing.Icon(stream))
                        {
                            _cachedIcon = Imaging.CreateBitmapSourceFromHIcon(
                                icon.Handle,
                                Int32Rect.Empty,
                                BitmapSizeOptions.FromEmptyOptions());

                            // フリーズしてスレッド間共有可能にする
                            if (_cachedIcon.CanFreeze) _cachedIcon.Freeze();
                        }
                    }
                }
            }
            catch { }
        }

        private void LoadSettings()
        {
            var settings = AppSettings.Load();
            this.SettingOffsetX = settings.OffsetX;
            this.SettingOffsetY = settings.OffsetY;
            this.SettingOpacity = settings.Opacity;
            this.SettingUpdateInterval = settings.UpdateInterval;
            this.SettingHideWhenCursorHidden = settings.HideWhenCursorHidden;
            this.SettingTextColor = Color.FromRgb(settings.TextR, settings.TextG, settings.TextB);
            this.SettingBackgroundColor = Color.FromRgb(settings.BgR, settings.BgG, settings.BgB);
        }

        private void SaveSettings()
        {
            var settings = new AppSettings
            {
                OffsetX = this.SettingOffsetX,
                OffsetY = this.SettingOffsetY,
                Opacity = this.SettingOpacity,
                UpdateInterval = this.SettingUpdateInterval,
                HideWhenCursorHidden = this.SettingHideWhenCursorHidden,
                TextR = this.SettingTextColor.R,
                TextG = this.SettingTextColor.G,
                TextB = this.SettingTextColor.B,
                BgR = this.SettingBackgroundColor.R,
                BgG = this.SettingBackgroundColor.G,
                BgB = this.SettingBackgroundColor.B
            };
            AppSettings.Save(settings);
        }

        private void UpdateBackgroundBrush()
        {
            if (MainBorder != null)
            {
                byte alpha = (byte)(SettingOpacity * 255 / 100);
                Color color = Color.FromArgb(alpha, SettingBackgroundColor.R, SettingBackgroundColor.G, SettingBackgroundColor.B);
                MainBorder.Background = new SolidColorBrush(color);
            }
        }

        public void ResetSettings()
        {
            SettingOffsetX = 10;
            SettingOffsetY = 10;
            SettingUpdateInterval = 10;
            SettingHideWhenCursorHidden = true;
            SettingTextColor = Colors.White;
            SettingBackgroundColor = Colors.Black;
            SettingOpacity = 67;
        }

        #endregion

        #region NotifyIcon & Settings

        private void InitializeNotifyIcon()
        {
            _notifyIcon = new Forms.NotifyIcon();
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("Imel.Imel.ico"))
                {
                    if (stream != null) _notifyIcon.Icon = new Drawing.Icon(stream);
                    else _notifyIcon.Icon = Drawing.SystemIcons.Application;
                }
            }
            catch
            {
                _notifyIcon.Icon = Drawing.SystemIcons.Application;
            }

            _notifyIcon.Text = "Imel (IME Indicator)";
            _notifyIcon.Visible = true;

            var contextMenu = new Forms.ContextMenuStrip();
            var settingsItem = new Forms.ToolStripMenuItem("設定...");
            settingsItem.Click += (s, e) => OpenSettings();
            var exitItem = new Forms.ToolStripMenuItem("終了");
            exitItem.Click += (s, e) => ExitApp();

            contextMenu.Items.Add(settingsItem);
            contextMenu.Items.Add(new Forms.ToolStripSeparator());
            contextMenu.Items.Add(exitItem);
            _notifyIcon.ContextMenuStrip = contextMenu;
            _notifyIcon.DoubleClick += (s, e) => OpenSettings();
        }

        private void OpenSettings()
        {
            foreach (Window w in Application.Current.Windows)
            {
                if (w is SettingsWindow)
                {
                    w.Activate();
                    return;
                }
            }
            var settingsWindow = new SettingsWindow(this);

            // キャッシュ済みのアイコンを設定
            if (_cachedIcon != null)
            {
                settingsWindow.Icon = _cachedIcon;
            }

            settingsWindow.Show();
        }

        private void ExitApp()
        {
            SaveSettings();
            _notifyIcon.Dispose();
            Application.Current.Shutdown();
        }

        #endregion

        #region Core Logic (Timer Loop)

        private void Timer_Tick(object? sender, EventArgs e)
        {
            try { ProcessUpdate(); }
            catch { this.Visibility = Visibility.Hidden; }
        }

        /// <summary>
        /// 定期更新処理のメインフロー
        /// </summary>
        private void ProcessUpdate()
        {
            // 1. カーソル可視チェック
            if (SettingHideWhenCursorHidden && !IsCursorVisible())
            {
                this.Visibility = Visibility.Hidden;
                return;
            }

            // 2. アクティブウィンドウ取得
            IntPtr hwndForeground = GetForegroundWindow();
            if (hwndForeground == IntPtr.Zero)
            {
                this.Visibility = Visibility.Hidden;
                return;
            }

            uint processId;
            uint threadId = GetWindowThreadProcessId(hwndForeground, out processId);

            // 3. IME状態チェック (間引き実行)
            if ((DateTime.Now - _lastImeCheckTime).TotalMilliseconds >= ImeCheckInterval)
            {
                CheckImeStatus(hwndForeground, threadId);
                _lastImeCheckTime = DateTime.Now;
            }

            // 4. 位置更新 (毎フレーム実行して滑らかに追従)
            if (this.Visibility == Visibility.Visible)
            {
                UpdatePosition();
            }

            // 5. メモリ解放 (定期的実行)
            if ((DateTime.Now - _lastMemoryCleanupTime).TotalMilliseconds >= MemoryCleanupInterval)
            {
                MinimizeFootprint();
                _lastMemoryCleanupTime = DateTime.Now;
            }
        }

        /// <summary>
        /// カーソルが表示されているかOSに問い合わせます。
        /// </summary>
        private bool IsCursorVisible()
        {
            var info = new CURSORINFO();
            info.cbSize = Marshal.SizeOf(info);
            if (GetCursorInfo(ref info))
            {
                return info.flags == CURSOR_SHOWING;
            }
            return true; // 取得失敗時は表示とみなす
        }

        private void CheckImeStatus(IntPtr hwndForeground, uint threadId)
        {
            uint currentThreadId = GetCurrentThreadId();
            bool attached = false;

            // 入力コンテキスト取得のためスレッドアタッチ
            if (threadId != currentThreadId) attached = AttachThreadInput(currentThreadId, threadId, true);

            try
            {
                IntPtr hwndFocus = GetFocus();
                if (hwndFocus == IntPtr.Zero) hwndFocus = hwndForeground;

                bool isImeOpen = false;
                int conversionMode = 0;
                bool statusRetrieved = false;

                // 4-A. ImmGetContext (レガシー)
                IntPtr hImc = ImmGetContext(hwndFocus);
                if (hImc != IntPtr.Zero)
                {
                    try
                    {
                        isImeOpen = ImmGetOpenStatus(hImc);
                        int sentenceMode = 0;
                        if (ImmGetConversionStatus(hImc, out conversionMode, out sentenceMode))
                        {
                            statusRetrieved = true;
                        }
                    }
                    finally { ImmReleaseContext(hwndFocus, hImc); }
                }

                // 4-B. WM_IME_CONTROL (モダンアプリ対策)
                if (!statusRetrieved)
                {
                    IntPtr hImeWnd = ImmGetDefaultIMEWnd(hwndFocus);
                    if (hImeWnd != IntPtr.Zero)
                    {
                        IntPtr openStatus = SendMessage(hImeWnd, WM_IME_CONTROL, (IntPtr)IMC_GETOPENSTATUS, IntPtr.Zero);
                        isImeOpen = (openStatus.ToInt32() != 0);
                        if (isImeOpen)
                        {
                            IntPtr convMode = SendMessage(hImeWnd, WM_IME_CONTROL, (IntPtr)IMC_GETCONVERSIONMODE, IntPtr.Zero);
                            conversionMode = convMode.ToInt32();
                        }
                        statusRetrieved = true;
                    }
                }

                // --- 表示テキストの決定ロジック ---
                string statusText = "_A"; // デフォルトは半角英数扱い

                if (statusRetrieved && isImeOpen)
                {
                    if ((conversionMode & IME_CMODE_NATIVE) != 0)
                    {
                        // 日本語入力モード
                        if ((conversionMode & IME_CMODE_KATAKANA) != 0)
                        {
                            // カタカナモード
                            statusText = (conversionMode & IME_CMODE_FULLSHAPE) != 0 ? "カ" : "_ｶ";
                        }
                        else
                        {
                            // ひらがなモード
                            statusText = "あ";
                        }
                    }
                    else
                    {
                        // 英数モード（IME ON）
                        if ((conversionMode & IME_CMODE_FULLSHAPE) != 0)
                        {
                            statusText = "Ａ"; // 全角英数
                        }
                        else
                        {
                            statusText = "_A"; // 半角英数
                        }
                    }
                }
                else
                {
                    // IME OFF
                    statusText = "_A";
                }

                // 【最適化】値が変わったときのみUI更新
                if (ImeStatusText.Text != statusText)
                {
                    ImeStatusText.Text = statusText;
                }

                this.Visibility = Visibility.Visible;
            }
            finally
            {
                if (attached) AttachThreadInput(currentThreadId, threadId, false);
            }
        }

        // マウス位置に追従してウィンドウを移動
        private void UpdatePosition()
        {
            GetCursorPos(out POINT mousePt);
            // キャッシュ済みDPIと設定されたオフセットを使用
            this.Left = (mousePt.X / _dpiX) + SettingOffsetX + 5;
            this.Top = (mousePt.Y / _dpiY) + SettingOffsetY + 5;
        }

        public void MinimizeFootprint()
        {
            try
            {
                GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced);
                GC.WaitForPendingFinalizers();

                if (Environment.OSVersion.Platform == PlatformID.Win32NT)
                {
                    EmptyWorkingSet(_thisProcessHandle);
                }
            }
            catch { }
        }

        #endregion

        #region Win32 API Definitions

        [DllImport("user32.dll")] static extern bool GetCursorInfo(ref CURSORINFO pci);
        [DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr hObject);
        [DllImport("psapi.dll")] static extern int EmptyWorkingSet(IntPtr hwProc);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("user32.dll")] static extern IntPtr GetFocus();
        [DllImport("imm32.dll")] static extern IntPtr ImmGetContext(IntPtr hWnd);
        [DllImport("imm32.dll")] static extern bool ImmGetOpenStatus(IntPtr hIMC);
        [DllImport("imm32.dll")] static extern bool ImmGetConversionStatus(IntPtr hIMC, out int lpfdwConversion, out int lpfdwSentence);
        [DllImport("imm32.dll")] static extern bool ImmReleaseContext(IntPtr hWnd, IntPtr hIMC);
        [DllImport("imm32.dll")] static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")][return: MarshalAs(UnmanagedType.Bool)] static extern bool GetCursorPos(out POINT lpPoint);

        const int IME_CMODE_NATIVE = 0x0001;
        const int IME_CMODE_KATAKANA = 0x0002;
        const int IME_CMODE_FULLSHAPE = 0x0008;
        const int WM_IME_CONTROL = 0x0283;
        const int IMC_GETOPENSTATUS = 0x0005;
        const int IMC_GETCONVERSIONMODE = 0x0001;
        const int CURSOR_SHOWING = 0x00000001;

        [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X; public int Y; }

        [StructLayout(LayoutKind.Sequential)]
        public struct CURSORINFO
        {
            public Int32 cbSize;
            public Int32 flags;
            public IntPtr hCursor;
            public POINT ptScreenPos;
        }

        #endregion
    }
}