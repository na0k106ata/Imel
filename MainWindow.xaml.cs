using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Drawing = System.Drawing;
using Forms = System.Windows.Forms;

namespace Imel
{
    /// <summary>
    /// メインウィンドウ（IMEインジケーター）のロジック。
    /// Win32 APIを使用してIME状態を監視し、マウスカーソル付近に表示します。
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Fields

        private DispatcherTimer _timer;
        private Forms.NotifyIcon _notifyIcon = null!;

        // IME状態の取得は比較的重いため、カーソル追従とは別に頻度を制限します。
        private DateTime _lastImeCheckTime = DateTime.MinValue;
        private const double ImeCheckInterval = 100.0;
        private bool _isImeCheckRunning = false;

        private double _dpiX = 1.0;
        private double _dpiY = 1.0;

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

        private int _settingUpdateInterval = 16;
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
            // 透明な小ウィンドウの描画負荷を安定させるため、ソフトウェアレンダリングを使用します。
            RenderOptions.ProcessRenderMode = RenderMode.SoftwareOnly;

            InitializeComponent();

            LoadSettings();

            UpdateBackgroundBrush();
            if (ImeStatusText != null)
            {
                ImeStatusText.Foreground = new SolidColorBrush(SettingTextColor);
            }

            InitializeNotifyIcon();

            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromMilliseconds(SettingUpdateInterval);
            _timer.Tick += Timer_Tick;

            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // 初回表示時のちらつきを避けるため、起動直後は画面外に置きます。
            Left = -100;
            Top = -100;

            var helper = new WindowInteropHelper(this);
            int exStyle = GetWindowLong(helper.Handle, GWL_EXSTYLE);

            // Alt+Tabやタスクバーに出ないツールウィンドウとして扱います。
            SetWindowLong(helper.Handle, GWL_EXSTYLE, exStyle | WS_EX_TOOLWINDOW);

            var source = PresentationSource.FromVisual(this);
            if (source?.CompositionTarget != null)
            {
                _dpiX = source.CompositionTarget.TransformToDevice.M11;
                _dpiY = source.CompositionTarget.TransformToDevice.M22;
            }

            _timer.Start();
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

        private ImageSource? CreateAppIconImageSource()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using var stream = assembly.GetManifestResourceStream("Imel.Imel.ico");
                if (stream == null) return null;

                using var icon = new Drawing.Icon(stream);
                var imageSource = Imaging.CreateBitmapSourceFromHIcon(
                    icon.Handle,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

                if (imageSource.CanFreeze) imageSource.Freeze();
                return imageSource;
            }
            catch
            {
                return null;
            }
        }

        private void LoadSettings()
        {
            var settings = AppSettings.Load();
            SettingOffsetX = settings.OffsetX;
            SettingOffsetY = settings.OffsetY;
            SettingOpacity = settings.Opacity;
            SettingUpdateInterval = settings.UpdateInterval;
            SettingHideWhenCursorHidden = settings.HideWhenCursorHidden;
            SettingScale = settings.Scale;

            SettingTextColor = Color.FromRgb(settings.TextR, settings.TextG, settings.TextB);
            SettingBackgroundColor = Color.FromRgb(settings.BgR, settings.BgG, settings.BgB);
        }

        private void SaveSettings()
        {
            var settings = new AppSettings
            {
                OffsetX = SettingOffsetX,
                OffsetY = SettingOffsetY,
                Opacity = SettingOpacity,
                UpdateInterval = SettingUpdateInterval,
                HideWhenCursorHidden = SettingHideWhenCursorHidden,
                Scale = SettingScale,
                TextR = SettingTextColor.R,
                TextG = SettingTextColor.G,
                TextB = SettingTextColor.B,
                BgR = SettingBackgroundColor.R,
                BgG = SettingBackgroundColor.G,
                BgB = SettingBackgroundColor.B
            };
            AppSettings.Save(settings);
        }

        private void UpdateWindowSize()
        {
            Width = BaseSize * SettingScale;
            Height = BaseSize * SettingScale;

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
                var color = Color.FromArgb(alpha, SettingBackgroundColor.R, SettingBackgroundColor.G, SettingBackgroundColor.B);
                MainBorder.Background = new SolidColorBrush(color);
            }
        }

        public void ResetSettings()
        {
            SettingOffsetX = 10;
            SettingOffsetY = 10;
            SettingUpdateInterval = 16;
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
                using var stream = assembly.GetManifestResourceStream("Imel.Imel.ico");
                _notifyIcon.Icon = stream != null ? new Drawing.Icon(stream) : Drawing.SystemIcons.Application;
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
            var appIcon = CreateAppIconImageSource();
            if (appIcon != null) settingsWindow.Icon = appIcon;
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
            try
            {
                ProcessUpdate();
            }
            catch
            {
                Visibility = Visibility.Hidden;
            }
        }

        private void ProcessUpdate()
        {
            // カーソル追従はUIスレッドで軽く保ち、IME問い合わせだけを別スレッドへ逃がします。
            if (Visibility == Visibility.Visible)
            {
                UpdatePosition();
            }

            var now = DateTime.UtcNow;
            if (!_isImeCheckRunning && (now - _lastImeCheckTime).TotalMilliseconds >= ImeCheckInterval)
            {
                _lastImeCheckTime = now;
                _ = CheckImeStatusAsync();
            }
        }

        private void SetIndicatorVisibility(Visibility visibility)
        {
            if (Visibility != visibility)
            {
                Visibility = visibility;
            }
        }

        private bool IsCursorVisible()
        {
            var info = new CURSORINFO { cbSize = Marshal.SizeOf<CURSORINFO>() };
            if (GetCursorInfo(ref info))
            {
                return info.flags == CURSOR_SHOWING;
            }

            return true;
        }

        private async Task CheckImeStatusAsync()
        {
            _isImeCheckRunning = true;

            try
            {
                if (SettingHideWhenCursorHidden && !IsCursorVisible())
                {
                    SetIndicatorVisibility(Visibility.Hidden);
                    return;
                }

                IntPtr hwndForeground = GetForegroundWindow();
                if (hwndForeground == IntPtr.Zero)
                {
                    SetIndicatorVisibility(Visibility.Hidden);
                    return;
                }

                uint processId;
                uint threadId = GetWindowThreadProcessId(hwndForeground, out processId);
                string statusText = await Task.Run(() => GetImeStatusText(threadId));

                if (ImeStatusText.Text != statusText)
                {
                    ImeStatusText.Text = statusText;
                }

                SetIndicatorVisibility(Visibility.Visible);
            }
            catch
            {
                SetIndicatorVisibility(Visibility.Hidden);
            }
            finally
            {
                _isImeCheckRunning = false;
            }
        }

        /// <summary>
        /// 指定スレッドのIME状態を取得し、表示用テキストへ変換します。
        /// AttachThreadInputは使わず、タイムアウト付きメッセージ送信でハングの巻き添えを避けます。
        /// </summary>
        private string GetImeStatusText(uint threadId)
        {
            bool statusRetrieved = false;
            bool isImeOpen = false;
            int conversionMode = 0;

            IntPtr hwndTarget = IntPtr.Zero;

            var guiInfo = new GUITHREADINFO { cbSize = Marshal.SizeOf<GUITHREADINFO>() };
            if (GetGUIThreadInfo(threadId, ref guiInfo))
            {
                hwndTarget = guiInfo.hwndFocus;
            }

            if (hwndTarget == IntPtr.Zero)
            {
                hwndTarget = GetForegroundWindow();
            }

            if (hwndTarget != IntPtr.Zero)
            {
                IntPtr hImeWnd = ImmGetDefaultIMEWnd(hwndTarget);
                if (hImeWnd != IntPtr.Zero)
                {
                    IntPtr retOpen = SendMessageTimeout(
                        hImeWnd,
                        WM_IME_CONTROL,
                        (IntPtr)IMC_GETOPENSTATUS,
                        IntPtr.Zero,
                        SMTO_ABORTIFHUNG,
                        200,
                        out IntPtr resultOpen);

                    if (retOpen != IntPtr.Zero)
                    {
                        isImeOpen = resultOpen.ToInt32() != 0;
                        if (isImeOpen)
                        {
                            IntPtr retConv = SendMessageTimeout(
                                hImeWnd,
                                WM_IME_CONTROL,
                                (IntPtr)IMC_GETCONVERSIONMODE,
                                IntPtr.Zero,
                                SMTO_ABORTIFHUNG,
                                200,
                                out IntPtr resultConv);

                            if (retConv != IntPtr.Zero)
                            {
                                conversionMode = resultConv.ToInt32();
                            }
                        }

                        statusRetrieved = true;
                    }
                }
            }

            if (!statusRetrieved || !isImeOpen)
            {
                return "_A";
            }

            if ((conversionMode & IME_CMODE_NATIVE) != 0)
            {
                if ((conversionMode & IME_CMODE_KATAKANA) != 0)
                {
                    return (conversionMode & IME_CMODE_FULLSHAPE) != 0 ? "カ" : "_ｶ";
                }

                return "あ";
            }

            return (conversionMode & IME_CMODE_FULLSHAPE) != 0 ? "Ａ" : "_A";
        }

        private void UpdatePosition()
        {
            GetCursorPos(out POINT mousePt);

            double newLeft = (mousePt.X / _dpiX) + SettingOffsetX + 5;
            double newTop = (mousePt.Y / _dpiY) + SettingOffsetY + 5;

            if (Math.Abs(Left - newLeft) > 0.1)
            {
                Left = newLeft;
            }

            if (Math.Abs(Top - newTop) > 0.1)
            {
                Top = newTop;
            }
        }

        #endregion

        #region Win32 API Definitions

        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")] static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOOLWINDOW = 0x00000080;

        [DllImport("user32.dll")] static extern bool GetCursorInfo(ref CURSORINFO pci);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")] static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);
        [DllImport("imm32.dll")] static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hWnd);

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
        const uint SMTO_ABORTIFHUNG = 0x0002;

        [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X; public int Y; }
        [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        public struct CURSORINFO
        {
            public int cbSize;
            public int flags;
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
