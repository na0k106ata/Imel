using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;

namespace Imel
{
    public partial class SettingsWindow : Window
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
            // スタートアップ
            StartupCheckBox.IsChecked = IsStartupEnabled();

            // カーソル連動
            HideCursorCheckBox.IsChecked = _mainWindow.SettingHideWhenCursorHidden;

            // 更新頻度
            IntervalSlider.Value = _mainWindow.SettingUpdateInterval;
            IntervalValueText.Text = _mainWindow.SettingUpdateInterval.ToString();

            // サイズ倍率 (New)
            ScaleSlider.Value = _mainWindow.SettingScale;
            ScaleValueText.Text = _mainWindow.SettingScale.ToString("F1");

            // 文字色
            SetRGBInputs(TextR, TextG, TextB, _mainWindow.SettingTextColor);
            UpdateComboFromColor(TextColorCombo, _mainWindow.SettingTextColor, true);

            // 背景色
            SetRGBInputs(BgR, BgG, BgB, _mainWindow.SettingBackgroundColor);
            UpdateComboFromColor(BgColorCombo, _mainWindow.SettingBackgroundColor, false);

            // 透過率
            BgOpacity.Text = _mainWindow.SettingOpacity.ToString();

            // 位置
            OffsetX.Text = _mainWindow.SettingOffsetX.ToString();
            OffsetY.Text = _mainWindow.SettingOffsetY.ToString();
        }

        // --- イベントハンドラ ---

        private void HideCursorCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            _mainWindow.SettingHideWhenCursorHidden = HideCursorCheckBox.IsChecked == true;
        }

        private void IntervalSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!_isInitialized) return;
            int val = (int)e.NewValue;
            _mainWindow.SettingUpdateInterval = val;
            if (IntervalValueText != null) IntervalValueText.Text = val.ToString();
        }

        // サイズ倍率変更 (New)
        private void ScaleSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!_isInitialized) return;
            double val = Math.Round(e.NewValue, 1);
            _mainWindow.SettingScale = val;
            if (ScaleValueText != null) ScaleValueText.Text = val.ToString("F1");
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

        private void StartupCheckBox_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(StartupRegistryKey, true);
                if (key == null) return;

                if (StartupCheckBox.IsChecked == true)
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
                MessageBox.Show($"設定の変更に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                StartupCheckBox.IsChecked = !StartupCheckBox.IsChecked;
            }
        }

        // --- 色・位置などの共通処理 ---

        private void SetRGBInputs(TextBox r, TextBox g, TextBox b, Color c)
        {
            r.Text = c.R.ToString();
            g.Text = c.G.ToString();
            b.Text = c.B.ToString();
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

        private void TextColorRGB_Changed(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            if (TryParseColor(TextR.Text, TextG.Text, TextB.Text, out Color c))
            {
                _mainWindow.SettingTextColor = c;
                _isInitialized = false;
                TextColorCombo.SelectedIndex = 5;
                _isInitialized = true;
            }
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

        private void BgColorRGB_Changed(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            if (TryParseColor(BgR.Text, BgG.Text, BgB.Text, out Color c))
            {
                _mainWindow.SettingBackgroundColor = c;
                _isInitialized = false;
                BgColorCombo.SelectedIndex = 5;
                _isInitialized = true;
            }
        }

        private void BgOpacity_Changed(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            if (int.TryParse(BgOpacity.Text, out int val)) _mainWindow.SettingOpacity = val;
        }

        private void Offset_Changed(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            if (int.TryParse(OffsetX.Text, out int x)) _mainWindow.SettingOffsetX = x;
            if (int.TryParse(OffsetY.Text, out int y)) _mainWindow.SettingOffsetY = y;
        }

        private bool TryParseColor(string rText, string gText, string bText, out Color color)
        {
            color = Colors.Black;
            if (byte.TryParse(rText, out byte r) &&
                byte.TryParse(gText, out byte g) &&
                byte.TryParse(bText, out byte b))
            {
                color = Color.FromRgb(r, g, b);
                return true;
            }
            return false;
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