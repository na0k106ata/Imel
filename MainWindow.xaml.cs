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
    // v1.0.3-dev: キー入力フリーズ対策のため、AttachThreadInputを廃止しSendMessageTimeoutを採用
    public partial class MainWindow : Window
    {
        #region Fields

        private DispatcherTimer _timer;
        private Forms.NotifyIcon _notifyIcon = null!;
        private ImageSource? _cachedIcon = null;

        private DateTime _lastImeCheckTime = DateTime.MinValue;
        private const double ImeCheckInterval = 100.0;

        private DateTime _lastMemoryCleanupTime = DateTime.MinValue;
        private const double MemoryCleanupInterval = 30000.0;

        private double _dpiX = 1.0;
        private double _dpiY = 1.0;
        private IntPtr _thisProcessHandle;

        private const double BaseSize = 24.0;
        private const double BaseFontSize = 13.0;

        #endregion

        #region Properties (Settings)

        public int SettingOffsetX { get; set; } = 10;
        public int SettingOffsetY { get; set; } = 10;
        public bool SettingHideWhenCursorHidden { get; set; } = true;

        private double _settingScale = 1.0;
        public double SettingScale
        {
            get => _settingScale;
            set
            {
                _settingScale = Math.Clamp(value, 0.5, 2.0);
                UpdateWindowSize();
            }
        }

        private int _settingUpdateInterval = 10;
        public int SettingUpdateInterval
        {
            get => _settingUpdateInterval;
            set
            {
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
            RenderOptions.ProcessRenderMode = RenderMode.SoftwareOnly;

            InitializeComponent();

            using (var proc = Process.GetCurrentProcess())
            {
                _thisProcessHandle = proc.Handle;
            }

            LoadSettings();

            UpdateBackgroundBrush();
            if (ImeStatusText != null)
                ImeStatusText.Foreground = new SolidColorBrush(SettingTextColor);

            InitializeNotifyIcon();
            CacheAppIcon();

            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromMilliseconds(SettingUpdateInterval);
            _timer.Tick += Timer_Tick;

            this.Loaded += MainWindow_Loaded;
            this.Closing += MainWindow_Closing;
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.Left = -100;
            this.Top = -100;

            var helper = new WindowInteropHelper(this);
            int exStyle = GetWindowLong(helper.Handle, GWL_EXSTYLE);
            // ツールウィンドウとして設定し、Alt+Tabなどに表示されないようにする
            SetWindowLong(helper.Handle, GWL_EXSTYLE, exStyle | WS_EX_TOOLWINDOW);

            var source = PresentationSource.FromVisual(this);
            if (source != null && source.CompositionTarget != null)
            {
                _dpiX = source.CompositionTarget.TransformToDevice.M11;
                _dpiY = source.CompositionTarget.TransformToDevice.M22;
            }

            _timer.Start();

            // 起動直後の安定化を待ってからメモリクリーンアップを実行
            await Task.Delay(2000);
            MinimizeFootprint();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveSettings();
            _notifyIcon.Dispose();
        }

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
            this.SettingScale = settings.Scale;

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
                Scale = this.SettingScale,

                TextR = this.SettingTextColor.R,
                TextG = this.SettingTextColor.G,
                TextB = this.SettingTextColor.B,
                BgR = this.SettingBackgroundColor.R,
                BgG = this.SettingBackgroundColor.G,
                BgB = this.SettingBackgroundColor.B
            };
            AppSettings.Save(settings);
        }

        private void UpdateWindowSize()
        {
            this.Width = BaseSize * SettingScale;
            this.Height = BaseSize * SettingScale;

            if (ImeStatusText != null)
            {
                ImeStatusText.FontSize = BaseFontSize * SettingScale;
            }
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
            SettingScale = 1.0;

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
            if (_cachedIcon != null) settingsWindow.Icon = _cachedIcon;
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

        private void ProcessUpdate()
        {
            if (SettingHideWhenCursorHidden && !IsCursorVisible())
            {
                this.Visibility = Visibility.Hidden;
                return;
            }

            IntPtr hwndForeground = GetForegroundWindow();
            if (hwndForeground == IntPtr.Zero)
            {
                this.Visibility = Visibility.Hidden;
                return;
            }

            uint processId;
            uint threadId = GetWindowThreadProcessId(hwndForeground, out processId);

            if ((DateTime.Now - _lastImeCheckTime).TotalMilliseconds >= ImeCheckInterval)
            {
                CheckImeStatus(threadId); // AttachThreadInputを使わないため引数を変更
                _lastImeCheckTime = DateTime.Now;
            }

            if (this.Visibility == Visibility.Visible)
            {
                UpdatePosition();
            }

            if ((DateTime.Now - _lastMemoryCleanupTime).TotalMilliseconds >= MemoryCleanupInterval)
            {
                MinimizeFootprint();
                _lastMemoryCleanupTime = DateTime.Now;
            }
        }

        private bool IsCursorVisible()
        {
            var info = new CURSORINFO();
            info.cbSize = Marshal.SizeOf(info);
            if (GetCursorInfo(ref info))
            {
                return info.flags == CURSOR_SHOWING;
            }
            return true;
        }

        // 修正: AttachThreadInputを廃止し、GetGUIThreadInfoとSendMessageTimeoutを使用
        private void CheckImeStatus(uint threadId)
        {
            bool statusRetrieved = false;
            bool isImeOpen = false;
            int conversionMode = 0;

            IntPtr hwndTarget = IntPtr.Zero;

            // 1. 指定スレッドのGUI情報を取得して、実際のフォーカスウィンドウ(キャレットがある場所)を特定する
            var guiInfo = new GUITHREADINFO();
            guiInfo.cbSize = Marshal.SizeOf(guiInfo);

            if (GetGUIThreadInfo(threadId, ref guiInfo))
            {
                hwndTarget = guiInfo.hwndFocus;
            }

            // フォーカスが取れない場合はフォアグラウンドウィンドウ自体をターゲットにする（フォールバック）
            if (hwndTarget == IntPtr.Zero)
            {
                hwndTarget = GetForegroundWindow();
            }

            if (hwndTarget != IntPtr.Zero)
            {
                // 2. ターゲットウィンドウに関連付けられたデフォルトIMEウィンドウを取得
                IntPtr hImeWnd = ImmGetDefaultIMEWnd(hwndTarget);
                if (hImeWnd != IntPtr.Zero)
                {
                    // 3. SendMessageTimeout で非同期(タイムアウト付き)にIME状態を問い合わせる
                    // フリーズ回避のため、SMTO_ABORTIFHUNGを使用し、タイムアウトを200msに設定

                    IntPtr resultOpen;
                    IntPtr resultConv;

                    // IMEが開いているか確認
                    IntPtr retOpen = SendMessageTimeout(
                        hImeWnd,
                        WM_IME_CONTROL,
                        (IntPtr)IMC_GETOPENSTATUS,
                        IntPtr.Zero,
                        SMTO_ABORTIFHUNG,
                        200,
                        out resultOpen);

                    if (retOpen != IntPtr.Zero) // 送信成功
                    {
                        isImeOpen = (resultOpen.ToInt32() != 0);
                        if (isImeOpen)
                        {
                            // 変換モードの取得
                            IntPtr retConv = SendMessageTimeout(
                                hImeWnd,
                                WM_IME_CONTROL,
                                (IntPtr)IMC_GETCONVERSIONMODE,
                                IntPtr.Zero,
                                SMTO_ABORTIFHUNG,
                                200,
                                out resultConv);

                            if (retConv != IntPtr.Zero)
                            {
                                conversionMode = resultConv.ToInt32();
                            }
                        }
                        statusRetrieved = true;
                    }
                }
            }

            string statusText = "_A";

            if (statusRetrieved && isImeOpen)
            {
                if ((conversionMode & IME_CMODE_NATIVE) != 0)
                {
                    if ((conversionMode & IME_CMODE_KATAKANA) != 0)
                    {
                        statusText = (conversionMode & IME_CMODE_FULLSHAPE) != 0 ? "カ" : "_ｶ";
                    }
                    else
                    {
                        statusText = "あ";
                    }
                }
                else
                {
                    if ((conversionMode & IME_CMODE_FULLSHAPE) != 0)
                    {
                        statusText = "Ａ";
                    }
                    else
                    {
                        statusText = "_A";
                    }
                }
            }
            else
            {
                statusText = "_A";
            }

            if (ImeStatusText.Text != statusText)
            {
                ImeStatusText.Text = statusText;
            }

            this.Visibility = Visibility.Visible;
        }

        private void UpdatePosition()
        {
            GetCursorPos(out POINT mousePt);
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

        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")] static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOOLWINDOW = 0x00000080;

        [DllImport("user32.dll")] static extern bool GetCursorInfo(ref CURSORINFO pci);
        [DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr hObject);
        [DllImport("psapi.dll")] static extern int EmptyWorkingSet(IntPtr hwProc);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // 変更: AttachThreadInput, GetFocus, ImmGetContext などを削除し、以下を追加

        [DllImport("user32.dll")] static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

        [DllImport("imm32.dll")] static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hWnd);

        // SendMessageTimeoutの定義追加
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern IntPtr SendMessageTimeout(
            IntPtr hWnd,
            uint Msg,
            IntPtr wParam,
            IntPtr lParam,
            uint fuFlags,
            uint uTimeout,
            out IntPtr lpdwResult);

        [DllImport("user32.dll")][return: MarshalAs(UnmanagedType.Bool)] static extern bool GetCursorPos(out POINT lpPoint);

        const int IME_CMODE_NATIVE = 0x0001;
        const int IME_CMODE_KATAKANA = 0x0002;
        const int IME_CMODE_FULLSHAPE = 0x0008;
        const int WM_IME_CONTROL = 0x0283;
        const int IMC_GETOPENSTATUS = 0x0005;
        const int IMC_GETCONVERSIONMODE = 0x0001;
        const int CURSOR_SHOWING = 0x00000001;

        // SendMessageTimeout flags
        const uint SMTO_ABORTIFHUNG = 0x0002;

        [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X; public int Y; }

        [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        public struct CURSORINFO
        {
            public Int32 cbSize;
            public Int32 flags;
            public IntPtr hCursor;
            public POINT ptScreenPos;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct GUITHREADINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        }

        #endregion
    }
}