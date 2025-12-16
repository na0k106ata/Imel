using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using Wpf.Ui.Controls; // FluentWindow, NumberBox等のため

namespace Imel
{
    // partial クラスの基底クラスを FluentWindow に統一
    public partial class SettingsWindow : FluentWindow
    {
        private MainWindow _mainWindow;
        private bool _isInitialized = false;

        private const string StartupRegistryKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string AppName = "Imel";

        public SettingsWindow(MainWindow mainWindow)
        {
            InitializeComponent();
            _mainWindow = mainWindow;

            LoadCurrentSettings();
            _isInitialized = true;
        }

        private void LoadCurrentSettings()
        {
            // ToggleSwitchへ反映
            StartupSwitch.IsChecked = IsStartupEnabled();
            HideCursorSwitch.IsChecked = _mainWindow.SettingHideWhenCursorHidden;

            // スライダー
            IntervalSlider.Value = _mainWindow.SettingUpdateInterval;
            IntervalValueText.Text = $"{_mainWindow.SettingUpdateInterval} ms";

            ScaleSlider.Value = _mainWindow.SettingScale;
            ScaleValueText.Text = $"{_mainWindow.SettingScale:F1} x";

            // 色設定
            SetRGBInputs(TextR, TextG, TextB, _mainWindow.SettingTextColor);
            UpdateComboFromColor(TextColorCombo, _mainWindow.SettingTextColor, true);

            SetRGBInputs(BgR, BgG, BgB, _mainWindow.SettingBackgroundColor);
            UpdateComboFromColor(BgColorCombo, _mainWindow.SettingBackgroundColor, false);

            // NumberBoxへ反映
            BgOpacity.Value = _mainWindow.SettingOpacity;
            OffsetX.Value = _mainWindow.SettingOffsetX;
            OffsetY.Value = _mainWindow.SettingOffsetY;
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

        // --- スタートアップ設定 ---

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
                    string? path = Environment.ProcessPath;
                    if (!string.IsNullOrEmpty(path)) key.SetValue(AppName, $"\"{path}\"");
                }
                else
                {
                    key.DeleteValue(AppName, false);
                }
            }
            catch (Exception ex)
            {
                // エラー時はトグルを戻す (System.Windows.MessageBox を明示)
                System.Windows.MessageBox.Show($"設定の変更に失敗しました。\n{ex.Message}", "エラー", System.Windows.MessageBoxButton.OK, MessageBoxImage.Error);
                StartupSwitch.IsChecked = !StartupSwitch.IsChecked;
            }
        }

        // --- 色・位置などの共通処理 ---

        // 引数の型を Wpf.Ui.Controls.NumberBox に明示
        private void SetRGBInputs(Wpf.Ui.Controls.NumberBox r, Wpf.Ui.Controls.NumberBox g, Wpf.Ui.Controls.NumberBox b, Color c)
        {
            r.Value = c.R;
            g.Value = c.G;
            b.Value = c.B;
        }

        private void UpdateComboFromColor(System.Windows.Controls.ComboBox combo, Color c, bool isText)
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
                _isInitialized = false;
                SetRGBInputs(TextR, TextG, TextB, c);
                _mainWindow.SettingTextColor = c;
                _isInitialized = true;
            }
        }

        // NumberBoxのイベントハンドラ
        private void TextColorRGB_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            // NumberBox.Valueはdouble?型
            byte r = (byte)(TextR.Value ?? 0);
            byte g = (byte)(TextG.Value ?? 0);
            byte b = (byte)(TextB.Value ?? 0);

            Color c = Color.FromRgb(r, g, b);
            _mainWindow.SettingTextColor = c;

            _isInitialized = false;
            TextColorCombo.SelectedIndex = 5; // Custom
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
            BgColorCombo.SelectedIndex = 5; // Custom
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
            this.Close();
        }
    }
}