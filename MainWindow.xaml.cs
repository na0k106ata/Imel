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
        private readonly DispatcherTimer _imeTimer;
        private readonly DispatcherTimer _saveTimer;
        private bool _settingsLoaded;
        private bool _settingsDirty;
        public string? SettingsSaveError { get; private set; }
        public event EventHandler? SettingsSaveStatusChanged;
        private const double ImeCheckInterval = 100.0;
        private bool _isImeCheckRunning = false;
        private bool _isClosing;
        internal bool IsShuttingDown => _isClosing;

        private readonly record struct ImeTarget(IntPtr Foreground, IntPtr Focus, uint ThreadId, uint ProcessId);

        private IntPtr _windowHandle;

        private const double BaseSize = 24.0;
        private const double BaseFontSize = 13.0;

        #endregion

        #region Properties (Settings)

        private int _offsetX = 10, _offsetY = 10;
        private bool _hideWhenCursorHidden = true;
        private bool _flipAtScreenEdge;
        public bool SettingFlipAtScreenEdge
        {
            get => _flipAtScreenEdge;
            set { _flipAtScreenEdge = value; ScheduleSettingsSave(); }
        }
        public int SettingOffsetX { get => _offsetX; set { _offsetX = value; ScheduleSettingsSave(); } }
        public int SettingOffsetY { get => _offsetY; set { _offsetY = value; ScheduleSettingsSave(); } }
        public bool SettingHideWhenCursorHidden
        {
            get => _hideWhenCursorHidden;
            set { _hideWhenCursorHidden = value; ScheduleSettingsSave(); }
        }

        private double _settingScale = 1.0;
        public double SettingScale
        {
            get => _settingScale;
            set
            {
                _settingScale = Math.Clamp(value, 0.5, 2.0);
                UpdateWindowSize();
                ScheduleSettingsSave();
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
                ScheduleSettingsSave();
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
                ScheduleSettingsSave();
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
                ScheduleSettingsSave();
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
                ScheduleSettingsSave();
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

            // 設定画面を開いている間も追従が遅れにくいよう、位置更新タイマーは高めの優先度で動かします。
            _timer = new DispatcherTimer(DispatcherPriority.Send);
            _timer.Interval = TimeSpan.FromMilliseconds(SettingUpdateInterval);
            _timer.Tick += Timer_Tick;
            _imeTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(ImeCheckInterval)
            };
            _imeTimer.Tick += (_, _) =>
            {
                if (!_isClosing && !_isImeCheckRunning) _ = CheckImeStatusAsync();
            };
            _saveTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(500)
            };
            _saveTimer.Tick += (_, _) => FlushSettings();
            _settingsLoaded = true;

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

            _windowHandle = helper.Handle;
            SetIndicatorVisibility(Visibility.Hidden);
            _imeTimer.Start();
            _ = CheckImeStatusAsync();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            _isClosing = true;
            _timer.Stop();
            _imeTimer.Stop();
            _saveTimer.Stop();
            if (!FlushSettings())
            {
                System.Windows.MessageBox.Show("設定を保存できませんでした。\n" + SettingsSaveError,
                    "Imel", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            _notifyIcon.Dispose();
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
            SettingFlipAtScreenEdge = settings.FlipAtScreenEdge;
            SettingScale = settings.Scale;

            SettingTextColor = Color.FromRgb(settings.TextR, settings.TextG, settings.TextB);
            SettingBackgroundColor = Color.FromRgb(settings.BgR, settings.BgG, settings.BgB);
        }

        private void ScheduleSettingsSave()
        {
            if (!_settingsLoaded || _isClosing) return;
            _settingsDirty = true;
            _saveTimer.Stop();
            _saveTimer.Start();
        }

        public bool FlushSettings()
        {
            _saveTimer.Stop();
            if (!_settingsDirty) return true;
            return SaveSettings();
        }

        private bool SaveSettings()
        {
            var settings = new AppSettings
            {
                OffsetX = SettingOffsetX,
                OffsetY = SettingOffsetY,
                Opacity = SettingOpacity,
                UpdateInterval = SettingUpdateInterval,
                HideWhenCursorHidden = SettingHideWhenCursorHidden,
                FlipAtScreenEdge = SettingFlipAtScreenEdge,
                Scale = SettingScale,
                TextR = SettingTextColor.R,
                TextG = SettingTextColor.G,
                TextB = SettingTextColor.B,
                BgR = SettingBackgroundColor.R,
                BgG = SettingBackgroundColor.G,
                BgB = SettingBackgroundColor.B
            };
            bool saved = AppSettings.Save(settings, out string? error);
            SettingsSaveError = error;
            if (saved) _settingsDirty = false;
            SettingsSaveStatusChanged?.Invoke(this, EventArgs.Empty);
            return saved;
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
            SettingFlipAtScreenEdge = false;
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
                SetIndicatorVisibility(Visibility.Hidden);
            }
        }

        private void ProcessUpdate()
        {
            if (_isClosing) return;
            if (Visibility == Visibility.Visible && !UpdatePosition())
                SetIndicatorVisibility(Visibility.Hidden);
        }

        private void SetIndicatorVisibility(Visibility visibility)
        {
            if (Visibility != visibility) Visibility = visibility;
            if (visibility == Visibility.Visible && !_isClosing)
                _timer.Start();
            else
                _timer.Stop();
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

                if (!TryGetImeTarget(out var target))
                {
                    SetIndicatorVisibility(Visibility.Hidden);
                    return;
                }

                string? statusText = await Task.Run(() =>
                    ImeStatusReader.Read(target.Focus, ImmGetDefaultIMEWnd, QueryImeControl));

                if (_isClosing) return;

                // await 中にアプリや同じアプリ内の入力先が変わった結果は使わない。
                if (!TryGetImeTarget(out var currentTarget) || currentTarget != target)
                {
                    SetIndicatorVisibility(Visibility.Hidden);
                    return;
                }

                if (statusText == null || (SettingHideWhenCursorHidden && !IsCursorVisible()))
                {
                    SetIndicatorVisibility(Visibility.Hidden);
                    return;
                }

                if (ImeStatusText.Text != statusText)
                {
                    ImeStatusText.Text = statusText;
                }

                // 非表示中にカーソルが移動していても、古い場所に表示しない。
                SetIndicatorVisibility(UpdatePosition() ? Visibility.Visible : Visibility.Hidden);
            }
            catch
            {
                if (!_isClosing) SetIndicatorVisibility(Visibility.Hidden);
            }
            finally
            {
                _isImeCheckRunning = false;
            }
        }

        private static bool TryGetImeTarget(out ImeTarget target)
        {
            target = default;
            IntPtr foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero) return false;

            uint threadId = GetWindowThreadProcessId(foreground, out uint processId);
            if (threadId == 0) return false;

            var guiInfo = new GUITHREADINFO { cbSize = Marshal.SizeOf<GUITHREADINFO>() };
            if (!GetGUIThreadInfo(threadId, ref guiInfo)) return false;
            if (GetForegroundWindow() != foreground) return false;

            // フォーカスなしの場合も、取得済みの前面ウィンドウに対象を固定する。
            IntPtr focus = guiInfo.hwndFocus != IntPtr.Zero ? guiInfo.hwndFocus : foreground;
            target = new ImeTarget(foreground, focus, threadId, processId);
            return true;
        }

        private static bool QueryImeControl(IntPtr imeWindow, int command, out IntPtr result)
        {
            return SendMessageTimeout(imeWindow, WM_IME_CONTROL, (IntPtr)command,
                IntPtr.Zero, SMTO_ABORTIFHUNG, 200, out result) != IntPtr.Zero;
        }

        private bool UpdatePosition()
        {
            if (!GetCursorPos(out POINT mousePt)) return false;

            return ScreenPlacement.PlaceIndicator(_windowHandle, mousePt.X, mousePt.Y,
                Width, Height, SettingOffsetX + 5.0, SettingOffsetY + 5.0, SettingFlipAtScreenEdge);
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

        const int WM_IME_CONTROL = 0x0283;
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
