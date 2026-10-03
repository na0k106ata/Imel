using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;
using Wpf.Ui.Markup;

namespace Imel
{
    /// <summary>
    /// 設定画面のロジック。
    /// メインウィンドウが保持している設定値を編集し、必要に応じて即時反映します。
    /// </summary>
    public partial class SettingsWindow : FluentWindow
    {
        private readonly MainWindow _mainWindow;
        private bool _isInitialized = false;

        public SettingsWindow(MainWindow mainWindow)
        {
            InitializeComponent();
            _mainWindow = mainWindow;

            SourceInitialized += SettingsWindow_SourceInitialized;
            Loaded += (_, _) => ScreenPlacement.FitSettingsWindow(this);

            LoadCurrentSettings();
            _isInitialized = true;
            _mainWindow.SettingsSaveStatusChanged += SettingsSaveStatusChanged;
            Closing += SettingsWindow_Closing;
            Closed += (_, _) => _mainWindow.SettingsSaveStatusChanged -= SettingsSaveStatusChanged;
            SettingsSaveStatusChanged(this, EventArgs.Empty);
        }


        private void SettingsWindow_SourceInitialized(object? sender, EventArgs e)
        {
            ApplyCurrentSystemTheme();
            ScreenPlacement.FitSettingsWindow(this);
        }

        private void SettingsSaveStatusChanged(object? sender, EventArgs e)
        {
            SaveStatusText.Text = _mainWindow.SettingsSaveError == null
                ? "" : "設定を保存できません: " + _mainWindow.SettingsSaveError;
            SaveStatusText.Visibility = _mainWindow.SettingsSaveError == null
                ? Visibility.Collapsed : Visibility.Visible;
        }

        private void SettingsWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!_mainWindow.FlushSettings() && !_mainWindow.IsShuttingDown)
            {
                // 保存できない環境でも画面を閉じられるようにする。未保存の変更はアプリ終了時に再度保存を試みる。
                var result = System.Windows.MessageBox.Show("設定を保存できませんでした。保存先を確認してください。\n" +
                    _mainWindow.SettingsSaveError + "\n\n保存せずに設定画面を閉じますか？",
                    "Imel", System.Windows.MessageBoxButton.YesNo, MessageBoxImage.Warning);
                e.Cancel = result != System.Windows.MessageBoxResult.Yes;
            }
        }

        protected override void OnDpiChanged(DpiScale oldDpi, DpiScale newDpi)
        {
            base.OnDpiChanged(oldDpi, newDpi);
            Dispatcher.BeginInvoke(new Action(() => ScreenPlacement.FitSettingsWindow(this)));
        }

        private void ApplyCurrentSystemTheme()
        {
            var systemTheme = ApplicationThemeManager.GetSystemTheme();
            var applicationTheme = systemTheme switch
            {
                SystemTheme.Dark or SystemTheme.Glow or SystemTheme.CapturedMotion => ApplicationTheme.Dark,
                SystemTheme.HCBlack or SystemTheme.HCWhite or SystemTheme.HC1 or SystemTheme.HC2 => ApplicationTheme.HighContrast,
                _ => ApplicationTheme.Light
            };

            for (int i = Resources.MergedDictionaries.Count - 1; i >= 0; i--)
            {
                if (Resources.MergedDictionaries[i] is ThemesDictionary)
                {
                    Resources.MergedDictionaries.RemoveAt(i);
                }
            }

            Resources.MergedDictionaries.Insert(0, new ThemesDictionary { Theme = applicationTheme });
            WindowBackgroundManager.UpdateBackground(this, applicationTheme, WindowBackdropType.Mica);
        }

        private void LoadCurrentSettings()
        {
            StartupSwitch.IsChecked = StartupRegistration.IsEnabled();
            HideCursorSwitch.IsChecked = _mainWindow.SettingHideWhenCursorHidden;
            FlipAtScreenEdgeSwitch.IsChecked = _mainWindow.SettingFlipAtScreenEdge;

            IntervalSlider.Value = _mainWindow.SettingUpdateInterval;
            IntervalValueText.Text = $"{_mainWindow.SettingUpdateInterval} ms";

            ScaleSlider.Value = _mainWindow.SettingScale;
            ScaleValueText.Text = $"{_mainWindow.SettingScale:F1} x";

            SetRGBInputs(TextR, TextG, TextB, _mainWindow.SettingTextColor);
            UpdateComboFromColor(TextColorCombo, _mainWindow.SettingTextColor, true);

            SetRGBInputs(BgR, BgG, BgB, _mainWindow.SettingBackgroundColor);
            UpdateComboFromColor(BgColorCombo, _mainWindow.SettingBackgroundColor, false);

            BgOpacity.Value = _mainWindow.SettingOpacity;
            OffsetX.Value = _mainWindow.SettingOffsetX;
            OffsetY.Value = _mainWindow.SettingOffsetY;
        }

        private void HideCursorSwitch_Click(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            _mainWindow.SettingHideWhenCursorHidden = HideCursorSwitch.IsChecked == true;
        }

        private void FlipAtScreenEdgeSwitch_Click(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            _mainWindow.SettingFlipAtScreenEdge = FlipAtScreenEdgeSwitch.IsChecked == true;
        }

        private void IntervalSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!_isInitialized) return;
            int val = (int)e.NewValue;
            _mainWindow.SettingUpdateInterval = val;
            if (IntervalValueText != null) IntervalValueText.Text = $"{val} ms";
        }

        private void ScaleSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!_isInitialized) return;
            double val = Math.Round(e.NewValue, 1);
            _mainWindow.SettingScale = val;
            if (ScaleValueText != null) ScaleValueText.Text = $"{val:F1} x";
        }

        private void StartupSwitch_Click(object sender, RoutedEventArgs e)
        {
            bool enabled = StartupSwitch.IsChecked == true;
            try
            {
                StartupRegistration.SetEnabled(enabled);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"設定の変更に失敗しました。\n{ex.Message}",
                    "エラー",
                    System.Windows.MessageBoxButton.OK,
                    MessageBoxImage.Error);
                StartupSwitch.IsChecked = StartupRegistration.IsEnabled();
            }
        }

        private void SetRGBInputs(Wpf.Ui.Controls.NumberBox r, Wpf.Ui.Controls.NumberBox g, Wpf.Ui.Controls.NumberBox b, Color c)
        {
            r.Value = c.R;
            g.Value = c.G;
            b.Value = c.B;
        }

        // 確定前の入力値が範囲外でも、byte への変換で値が回り込まないようにする。
        private static byte ToColorByte(double? value)
        {
            double v = value ?? 0;
            return double.IsFinite(v) ? (byte)Math.Clamp(Math.Round(v), 0, 255) : (byte)0;
        }

        private void UpdateComboFromColor(ComboBox combo, Color c, bool isText)
        {
            if (isText)
            {
                if (c == Colors.White) combo.SelectedIndex = 0;
                else if (c == Colors.Black) combo.SelectedIndex = 1;
                else if (c == Colors.Red) combo.SelectedIndex = 2;
                else if (c == Colors.Blue) combo.SelectedIndex = 3;
                else if (c == Colors.Green) combo.SelectedIndex = 4;
                else combo.SelectedIndex = 5;
            }
            else
            {
                if (c == Colors.Black) combo.SelectedIndex = 0;
                else if (c == Colors.White) combo.SelectedIndex = 1;
                else if (c == Colors.Red) combo.SelectedIndex = 2;
                else if (c == Colors.Blue) combo.SelectedIndex = 3;
                else if (c == Colors.Green) combo.SelectedIndex = 4;
                else combo.SelectedIndex = 5;
            }
        }

        private void TextColorCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;
            Color c = Colors.White;
            bool isCustom = false;

            switch (TextColorCombo.SelectedIndex)
            {
                case 0: c = Colors.White; break;
                case 1: c = Colors.Black; break;
                case 2: c = Colors.Red; break;
                case 3: c = Colors.Blue; break;
                case 4: c = Colors.Green; break;
                default: isCustom = true; break;
            }

            if (!isCustom)
            {
                _isInitialized = false;
                SetRGBInputs(TextR, TextG, TextB, c);
                _mainWindow.SettingTextColor = c;
                _isInitialized = true;
            }
        }

        private void TextColorRGB_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;

            byte r = ToColorByte(TextR.Value);
            byte g = ToColorByte(TextG.Value);
            byte b = ToColorByte(TextB.Value);

            _mainWindow.SettingTextColor = Color.FromRgb(r, g, b);

            _isInitialized = false;
            TextColorCombo.SelectedIndex = 5;
            _isInitialized = true;
        }

        private void BgColorCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;
            Color c = Colors.Black;
            bool isCustom = false;

            switch (BgColorCombo.SelectedIndex)
            {
                case 0: c = Colors.Black; break;
                case 1: c = Colors.White; break;
                case 2: c = Colors.Red; break;
                case 3: c = Colors.Blue; break;
                case 4: c = Colors.Green; break;
                default: isCustom = true; break;
            }

            if (!isCustom)
            {
                _isInitialized = false;
                SetRGBInputs(BgR, BgG, BgB, c);
                _mainWindow.SettingBackgroundColor = c;
                _isInitialized = true;
            }
        }

        private void BgColorRGB_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;

            byte r = ToColorByte(BgR.Value);
            byte g = ToColorByte(BgG.Value);
            byte b = ToColorByte(BgB.Value);

            _mainWindow.SettingBackgroundColor = Color.FromRgb(r, g, b);

            _isInitialized = false;
            BgColorCombo.SelectedIndex = 5;
            _isInitialized = true;
        }

        private void BgOpacity_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            _mainWindow.SettingOpacity = (int)(BgOpacity.Value ?? 100);
        }

        private void Offset_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            _mainWindow.SettingOffsetX = (int)(OffsetX.Value ?? 0);
            _mainWindow.SettingOffsetY = (int)(OffsetY.Value ?? 0);
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            _isInitialized = false;
            _mainWindow.ResetSettings();
            LoadCurrentSettings();
            _isInitialized = true;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
