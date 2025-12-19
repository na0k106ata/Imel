using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using Wpf.Ui.Controls; // FluentWindow, NumberBox等のため

namespace Imel
{
    /// <summary>
    /// 設定ウィンドウ
    /// Wpf.Ui (Fluent Design) を使用してWindows 11ライクなUIを提供します。
    /// </summary>
    public partial class SettingsWindow : FluentWindow
    {
        private MainWindow _mainWindow;
        private bool _isInitialized = false;
        private AppSettings _currentSettings;

        private const string StartupRegistryKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string AppName = "Imel";

        public SettingsWindow(MainWindow mainWindow)
        {
            InitializeComponent();
            _mainWindow = mainWindow;

            // 設定を読み込んでおく（保存用）
            _currentSettings = AppSettings.Load();

            LoadCurrentSettings();
            _isInitialized = true;
        }

        /// <summary>
        /// MainWindowの現在の設定値をUIコントロールに反映させます
        /// </summary>
        private void LoadCurrentSettings()
        {
            // スタートアップ設定とカーソル連動設定の反映
            StartupSwitch.IsChecked = IsStartupEnabled();
            HideCursorSwitch.IsChecked = _mainWindow.SettingHideWhenCursorHidden;

            // テーマ設定の反映
            // Enumの値(Light=0, Dark=1, Auto=2) と コンボボックスの並び順(Auto=0, Light=1, Dark=2) をマッピング
            ThemeCombo.SelectedIndex = _currentSettings.Theme switch
            {
                AppTheme.Light => 1,
                AppTheme.Dark => 2,
                _ => 0 // Auto
            };

            // スライダー値の反映
            IntervalSlider.Value = _mainWindow.SettingUpdateInterval;
            IntervalValueText.Text = $"{_mainWindow.SettingUpdateInterval} ms";

            ScaleSlider.Value = _mainWindow.SettingScale;
            ScaleValueText.Text = $"{_mainWindow.SettingScale:F1} x";

            // 色設定の反映（プリセット判定含む）
            SetRGBInputs(TextR, TextG, TextB, _mainWindow.SettingTextColor);
            UpdateComboFromColor(TextColorCombo, _mainWindow.SettingTextColor, true);

            SetRGBInputs(BgR, BgG, BgB, _mainWindow.SettingBackgroundColor);
            UpdateComboFromColor(BgColorCombo, _mainWindow.SettingBackgroundColor, false);

            // 数値入力ボックスへの反映
            BgOpacity.Value = _mainWindow.SettingOpacity;
            OffsetX.Value = _mainWindow.SettingOffsetX;
            OffsetY.Value = _mainWindow.SettingOffsetY;
        }

        // --- テーマ設定 ---

        private void ThemeCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;

            AppTheme selectedTheme = AppTheme.Auto;
            switch (ThemeCombo.SelectedIndex)
            {
                case 0: selectedTheme = AppTheme.Auto; break;
                case 1: selectedTheme = AppTheme.Light; break;
                case 2: selectedTheme = AppTheme.Dark; break;
            }

            // テーマを即時適用
            App.ApplyTheme(selectedTheme);

            // 設定オブジェクトに保存
            _currentSettings.Theme = selectedTheme;
            AppSettings.Save(_currentSettings);
        }

        // --- イベントハンドラ ---

        private void HideCursorSwitch_Click(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            _mainWindow.SettingHideWhenCursorHidden = HideCursorSwitch.IsChecked == true;
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

        // --- スタートアップ設定 (レジストリ操作) ---

        private bool IsStartupEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(StartupRegistryKey, false);
                return key?.GetValue(AppName) != null;
            }
            catch { return false; }
        }

        private void StartupSwitch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(StartupRegistryKey, true);
                if (key == null) return;

                if (StartupSwitch.IsChecked == true)
                {
                    // 実行ファイルのパスをレジストリに登録
                    string? path = Environment.ProcessPath;
                    if (!string.IsNullOrEmpty(path)) key.SetValue(AppName, $"\"{path}\"");
                }
                else
                {
                    // レジストリから削除
                    key.DeleteValue(AppName, false);
                }
            }
            catch (Exception ex)
            {
                // 権限エラー等が発生した場合はユーザーに通知し、スイッチを元の状態に戻す
                System.Windows.MessageBox.Show($"設定の変更に失敗しました。\n{ex.Message}", "エラー", System.Windows.MessageBoxButton.OK, MessageBoxImage.Error);
                StartupSwitch.IsChecked = !StartupSwitch.IsChecked;
            }
        }

        // --- 色・位置などの共通処理 ---

        private void SetRGBInputs(Wpf.Ui.Controls.NumberBox r, Wpf.Ui.Controls.NumberBox g, Wpf.Ui.Controls.NumberBox b, Color c)
        {
            r.Value = c.R;
            g.Value = c.G;
            b.Value = c.B;
        }

        /// <summary>
        /// 現在の色がプリセットに含まれているか判定し、コンボボックスを選択状態にします
        /// </summary>
        private void UpdateComboFromColor(System.Windows.Controls.ComboBox combo, Color c, bool isText)
        {
            if (isText)
            {
                if (c == Colors.White) combo.SelectedIndex = 0;
                else if (c == Colors.Black) combo.SelectedIndex = 1;
                else if (c == Colors.Red) combo.SelectedIndex = 2;
                else if (c == Colors.Blue) combo.SelectedIndex = 3;
                else if (c == Colors.Green) combo.SelectedIndex = 4;
                else combo.SelectedIndex = 5; // Custom
            }
            else
            {
                if (c == Colors.Black) combo.SelectedIndex = 0;
                else if (c == Colors.White) combo.SelectedIndex = 1;
                else if (c == Colors.Red) combo.SelectedIndex = 2;
                else if (c == Colors.Blue) combo.SelectedIndex = 3;
                else if (c == Colors.Green) combo.SelectedIndex = 4;
                else combo.SelectedIndex = 5; // Custom
            }
        }

        private void TextColorCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
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
                // プリセット選択時はRGBボックスとメインウィンドウを更新
                _isInitialized = false;
                SetRGBInputs(TextR, TextG, TextB, c);
                _mainWindow.SettingTextColor = c;
                _isInitialized = true;
            }
        }

        private void TextColorRGB_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;

            byte r = (byte)(TextR.Value ?? 0);
            byte g = (byte)(TextG.Value ?? 0);
            byte b = (byte)(TextB.Value ?? 0);

            Color c = Color.FromRgb(r, g, b);
            _mainWindow.SettingTextColor = c;

            // 手動変更されたためコンボボックスをCustomに変更
            _isInitialized = false;
            TextColorCombo.SelectedIndex = 5;
            _isInitialized = true;
        }

        private void BgColorCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
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
            byte r = (byte)(BgR.Value ?? 0);
            byte g = (byte)(BgG.Value ?? 0);
            byte b = (byte)(BgB.Value ?? 0);

            Color c = Color.FromRgb(r, g, b);
            _mainWindow.SettingBackgroundColor = c;

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
            // テーマもリセットする場合はここで対応
            _currentSettings.Theme = AppTheme.Auto;
            ThemeCombo.SelectedIndex = 0;
            _isInitialized = true;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}