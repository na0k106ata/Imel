using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using Wpf.Ui.Controls; // FluentWindow, NumberBox等のため

namespace Imel
{
    /// <summary>
    /// 設定画面のロジック。
    /// Wpf.Ui (Fluent Design) を使用して設定UIを提供します。
    /// </summary>
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

        /// <summary>
        /// MainWindowの現在の設定値をUIコントロールに反映させます。
        /// </summary>
        private void LoadCurrentSettings()
        {
            // スタートアップ設定とカーソル連動設定の反映
            StartupSwitch.IsChecked = IsStartupEnabled();
            HideCursorSwitch.IsChecked = _mainWindow.SettingHideWhenCursorHidden;

            // 更新間隔スライダーの反映
            IntervalSlider.Value = _mainWindow.SettingUpdateInterval;
            IntervalValueText.Text = $"{_mainWindow.SettingUpdateInterval} ms";

            // スケールスライダーの反映
            ScaleSlider.Value = _mainWindow.SettingScale;
            ScaleValueText.Text = $"{_mainWindow.SettingScale:F1} x";

            // 文字色設定の反映（プリセット判定含む）
            SetRGBInputs(TextR, TextG, TextB, _mainWindow.SettingTextColor);
            UpdateComboFromColor(TextColorCombo, _mainWindow.SettingTextColor, true);

            // 背景色設定の反映（プリセット判定含む）
            SetRGBInputs(BgR, BgG, BgB, _mainWindow.SettingBackgroundColor);
            UpdateComboFromColor(BgColorCombo, _mainWindow.SettingBackgroundColor, false);

            // その他数値入力ボックスへの反映
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

        // --- スタートアップ設定 (レジストリ操作) ---

        /// <summary>
        /// レジストリを確認し、スタートアップに登録されているか判定します。
        /// </summary>
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
        /// 指定された色がプリセットに含まれているか判定し、コンボボックスの選択状態を更新します。
        /// </summary>
        private void UpdateComboFromColor(System.Windows.Controls.ComboBox combo, Color c, bool isText)
        {
            // プリセット色の定義 (インデックス順)
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