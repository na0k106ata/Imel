# Imel (IME Indicator)

**Imel** is a simple IME status indicator for Windows 11.
The name blends "IME" and "Mieru" (見える, Japanese for "visible"), reflecting the goal of making your input state visible at a glance.

It shows the current input mode (such as `_A` or `あ`) right next to the mouse cursor, so you can check it without looking down at the taskbar.

**Version: v1.0.15**

Project page: https://tabunugoku.github.io/Imel/

> 日本語版は下部の折りたたみ内にあります。 / A Japanese version is in the collapsible section at the bottom of this page.

## Features

* **At-a-glance display**: Follows the mouse cursor and shows the current input mode.
    * Distinguishes half-width alphanumeric (`_A`), full-width alphanumeric (`Ａ`), Hiragana (`あ`), half-width Katakana (`_ｶ`), and full-width Katakana (`カ`).
* **Lightweight**: Designed as a resident tool with low memory use and low rendering overhead.
* **Customizable**:
    * **Appearance**: Fine-tune size (scale), text and background colors (RGB), background opacity, and position offset (DIP, -200 to 200).
    * **Flip placement at screen edges**: Available under "Position" in the settings window. **OFF** by default, which keeps the normal offset position. When ON, the indicator flips to the left/top at the right/bottom edges so it stays inside the work area. When OFF, it may be clipped at screen edges.
* **Cursor-aware**:
    * When the OS hides the mouse cursor (for example while watching a video), the indicator hides too (can be turned ON/OFF in settings).
    * Supports launching at Windows startup. If you move `Imel.exe` to another folder after registering, the registration is updated the next time you launch it from the new location.
* **Tray-resident**: Stays out of your way while remaining one click from the settings window.
* **Update notifications**: Turn on "Check for updates" in the settings window to get a Windows notification when a new version is released. **OFF** by default (see "Checking for updates").
* **Automatic settings save**: Settings are saved about 500 ms after you stop making changes, and when the settings window is closed. If saving fails, the settings window shows an error.
* **Handling IME query failures**: If the input mode cannot be obtained, the indicator is hidden instead of wrongly showing alphanumeric input.

## Requirements

* **OS**: Windows 11
* **Runtime**: .NET 8.0 Desktop Runtime

## Installation and Usage

### Using the binary
Download the zip file from the Releases page, extract it, and run `Imel.exe`.

### Building from source
1.  Clone this repository.
2.  Open `Imel.csproj` in Visual Studio 2022.
3.  Set the solution configuration to `Release` and build.
4.  Run `Imel.exe` in the `bin/Release/net8.0-windows` folder.

Release builds also copy `README.md` into the same folder.

### Saving and updating settings

Settings are stored in `%APPDATA%\Imel\settings.json`. Existing settings files load as they are. If "Flip placement at screen edges" is not set, it is treated as OFF, and "Reset to defaults" also returns it to OFF.

When saving, Imel writes to a temporary file and then replaces the original, keeping the previous valid settings as `settings.json.bak`. If the main file is corrupt or missing, settings are loaded from the backup; if neither can be read, the defaults are used.

### Checking for updates

When "Check for updates" under "General" in the settings window is ON, Imel checks GitHub for the latest release right after you turn it on, at app startup, and every 24 hours. At startup, no check is made if fewer than 24 hours have passed since the last one. If a check fails, it is retried after 1 hour.

* If a newer version exists, a Windows notification is shown. Each version is announced only once. Clicking the notification opens the release page in your browser.
* While an update is available, "Download vX.Y.Z..." appears at the top of the tray menu.
* The settings window shows the time and result of the last check, and "Check now" runs a manual check.
* Updates are never downloaded or installed automatically. Download the zip from the Releases page and replace the files yourself.
* Checks connect to `api.github.com`. Only ordinary HTTPS connection information (such as your IP address) and a User-Agent containing the Imel version are sent. Nothing is sent while the setting is OFF.

## Changelog

### v1.0.15 (2026-10-06)

* **Settings window**: Switches, sliders, color pickers, and numeric fields now have names that screen readers and other assistive technologies can read aloud.
* **Internal cleanup**: No change in behavior.
    * Moved the Win32 API declarations and tray handling out of `MainWindow` into separate files.
    * Added diagnostic output for failed operations (visible when a debugger is attached).
* **Development**: Added `.editorconfig`. Added a test that checks the `Imel.csproj` version matches the README; there are now 81 regression tests.

### v1.0.14 (2026-10-06)

* **Text fixes**:
    * Removed extra spaces between Japanese text and strings like "Imel" and "vX.Y.Z" in update notifications and the tray menu.
    * Reworded the "Auto-hide" description in the settings window to "Hide while the mouse cursor is hidden (e.g. while watching videos)".

### v1.0.13 (2026-10-06)

* **Update check improvements**:
    * The last check result (new version, whether it failed) is now saved and carried over to the tray menu and settings window after a restart. If a check fails, it is retried after 1 hour even across restarts.
    * Checks no longer stall when the last-check time is in the future (for example after changing the PC clock).
* **Focus fix**: Showing the indicator no longer steals focus from the app you are typing in.
* **Settings window improvements**:
    * If the opacity, offset, or RGB fields are emptied or contain a non-numeric value, the previous value is kept (previously they changed to 100 or 0).
    * Widened the numeric fields and the settings window so entered values are readable.
* **Library update**: Updated WPF-UI from 4.1.0 to 4.3.0.
* **Development**: Added automated build and regression tests with GitHub Actions and dependency update checks with Dependabot. There are now 78 regression tests.

### v1.0.12 (2026-10-05)

* **Update check target changed**: The repository checked for updates is now `tabunugoku/Imel`, following the GitHub account rename. If you use v1.0.10 or v1.0.11, please update to v1.0.12.
* **Copyright holder changed**: The copyright holder in LICENSE and the app's version info is now `tabunugoku`.
* **Build change**: Distributed files no longer contain the folder path of the PC that built them.

### v1.0.11 (2026-10-05)

* **Update check**: No check is made at startup if fewer than 24 hours (1 hour after a failure) have passed since the last one.
* **Settings window improvements**: When an out-of-range value is entered for the position offset or opacity, the value actually applied is now reflected in the field.
* **Minor fix**: The tray menu is now disposed on exit.

### v1.0.10 (2026-10-03)

* **Update notifications added**: Added "Check for updates" to the settings window. When ON, Imel checks GitHub for new versions and notifies you via a Windows notification and the tray menu. OFF by default.
* **Regression tests**: Added tests for version comparison, release info parsing, and saving/default values of the update-check settings.

### v1.0.9 (2026-10-03)

* **Settings window improvements**:
    * The settings window can now be closed after confirmation even when settings cannot be saved. Unsaved changes are retried when the app exits.
    * The save-error text color now follows the settings window theme (light/dark).
    * RGB input is clamped to 0–255 before being applied, fixing colors changing with out-of-range values.
    * Limited the position offset range to -200 to 200 (DIP). Out-of-range values in existing settings files are clamped on load.
* **Fixed duplicate exit warning**: Fixed a problem where the warning could be shown twice when exiting the app while settings could not be saved.
* **Startup registration improvements**:
    * If the registered `Imel.exe` is missing, Imel re-registers the current `Imel.exe` on launch.
    * Registration now works even if the registry key does not exist, and if registration fails, the switch reverts to the actual registration state.
* **Development**: Added the regression test project to the solution, with tests for startup registration path parsing and offset ranges.

### v1.0.8 (2026-09-26)

* **IME display fixes**:
    * If getting the IME state or conversion mode fails, the indicator is now hidden instead of showing half-width alphanumeric.
    * After an asynchronous query, the foreground window and input target are rechecked, and the result is discarded if the target has changed.
    * The position is updated before re-showing, and display updates from asynchronous work after exit are suppressed.
* **Settings saving improvements**:
    * Added delayed saving after changes and saving when the settings window closes.
    * Added replace-on-save through a temporary file, recovery from backup, and display of save failures.
* **Less work while hidden**:
    * Separated the position-tracking timer from the IME state check timer.
    * While hidden, the position-tracking timer stops and the 100 ms state check continues.
* **Screen and DPI support**:
    * Switched placement to use physical coordinates and the window's DPI, and explicitly declared Per-Monitor V2.
    * Added the "Flip placement at screen edges" toggle. OFF by default.
    * Made the settings window resizable and added logic to fit it in the work area on first display and on DPI changes.
    * Corrected the label "透過率" (transparency) to "背景の不透明度" (background opacity).
* **Regression tests**: Added tests for IME queries, settings saving/recovery, placement calculation, and saving/default values of the flip setting.

### v1.0.7 (2026-05-27)
* **[Fix] Indicator appearance changing after opening the settings window**:
    * Fixed the settings window's light/dark theme leaking into the whole app and affecting the main indicator's background color and corner rendering.
    * Theme resources are now applied only to the settings window, separating the main indicator's appearance from it.
    * Adjusted the position-update timer priority so the indicator keeps up while the settings window is open.
* **[Fix] Multiple-instance prevention**:
    * Imel now uses a named Mutex so that only one instance runs.

### v1.0.6 (2026-05-27)
* **Lighter and cheaper while resident**:
    * Reviewed how often the IME state, foreground window, and cursor visibility are checked, reducing constantly running work.
    * Moved IME state queries to the background so cursor tracking is less likely to stall.
    * The position is no longer reset when the mouse hasn't moved, avoiding unnecessary WPF updates.
    * Lazy-loaded the settings window icon image to reduce resources loaded during normal residence.
    * Removed periodic forced GC and working-set trimming, leaving memory management to the .NET runtime.
* **Settings adjustments**:
    * Changed the default rendering interval from 10 ms to 16 ms for a lighter default that still tracks well.
    * Updated the recommended-value text for the rendering interval in the settings window.
    * The settings window colors now follow the Windows app theme (light/dark).

### v1.0.5 (2025-12-20)
* **[Fix]** The theme feature (dark mode, etc.) caused display problems, so it was removed and the behavior rolled back to v1.0.3.
* Fine-tuned the settings window layout.

### v1.0.4 (2025-12-20)
* **Settings window enhancements and redesign**:
    * **Theme setting**: The app's color scheme could be chosen from Light, Dark, or System.
    * **Mica material**: Applied the Mica effect to the settings window background to blend in with the Windows 11 desktop.
    * **Layout optimization**: Fixed the window to a compact size and rearranged the UI so every setting is reachable without scrolling.

### v1.0.3 (2025-12-18)
* **[Important] Improved keyboard input stability**:
    * Fixed a critical bug where, under certain conditions, keyboard input could temporarily stop being accepted (freeze).
    * Reworked the internal logic and removed `AttachThreadInput` (sharing thread input). It is replaced with an asynchronous approach combining `SendMessageTimeout` and `GetGUIThreadInfo`, so Imel and the system's input handling are no longer dragged down when the monitored application stops responding.

### v1.0.2 (2025-12-17)
* **Modernized settings window UI**: Redesigned the settings window's look and layout into a modern, polished interface.
    * **Better overview**: Optimized the window size so every setting is reachable without scrolling.
    * **Better usability**: Reworked the RGB color settings and the placement of items for more intuitive customization.
* **Other**: Optimized the project structure and updated version information.

## Development Environment

* Visual Studio 2022
* C# (WPF / .NET 8)

## Development and Testing

Run the automated tests with `dotnet run --project Tests/Imel.RegressionTests.csproj`. The test project is also included in `Imel.sln`, so it is built along with the solution.
When making changes, in addition to the build and the automated logic tests, please verify the following on a real machine:

* Launch and exit, multiple-instance prevention, and tray operations.
* Switching apps, the display for each IME input mode, and recovery after the cursor was hidden.
* Settings persisting across changes, closing the window, and restarting the app.
* Flip placement ON/OFF, default OFF, and reset to defaults.
* Mixed 100% / 150% / 200% DPI, monitors with negative coordinates, tracking at screen edges, and the settings window display.

Stopping timers while hidden is implemented, but the reduction in CPU usage and power consumption has not been measured.

## License

This software is released under the [MIT License](LICENSE).
See the LICENSE file for details.

---

<details>
<summary>日本語 (Japanese)</summary>

# Imel (IME Indicator)

**Imel (アイメル)** は、Windows 11 向けのシンプルなIME状態表示ツールです。
名前は「IME」と「見える（Mieru）」を組み合わせた造語で、入力状態が一目でわかるようにという意図が込められています。

マウスカーソルのすぐそばに、現在の入力モード（「_A」や「あ」など）を表示します。視線をタスクバーに移動させることなく、手元で入力状態を確認できます。

**バージョン: v1.0.15**

紹介ページ: https://tabunugoku.github.io/Imel/

## 特徴

* **直感的な表示**: マウスカーソルに追従して、現在の入力モードを表示します。
    * 半角英数 (`_A`)、全角英数 (`Ａ`)、ひらがな (`あ`)、半角カタカナ (`_ｶ`)、全角カタカナ (`カ`) を判別可能です。
* **省リソース設計**: 常駐ツールとして、メモリ使用量や描画負荷に配慮して設計されています。
* **カスタマイズ性**:
    * **表示カスタマイズ**: サイズ（倍率）、文字色・背景色（RGB指定可）、背景の不透明度、位置オフセット（DIP、-200～200）などを微調整可能。
    * **画面端で配置を反転**: 設定画面の「位置調整」から選択できます。既定は **OFF** で、通常のオフセット位置を維持します。ONにすると、右端・下端では左側・上側へ反転し、作業領域内に収めます。OFFでは画面端で表示が欠ける場合があります。
* **カーソル連動**:
    * 動画視聴時など、OS側でマウスカーソルが非表示になると、自動的にインジケーターも隠れます（設定でON/OFF可）。
    * Windows起動時の自動実行（スタートアップ）に対応。登録後に `Imel.exe` を別のフォルダへ移した場合は、移した先で起動したときに登録先を更新します。
* **タスクトレイ常駐**: 作業を妨げず、いつでも設定画面へアクセス可能。
* **更新の通知**: 設定画面の「更新の確認」をONにすると、新しいバージョンが公開されたときにWindowsの通知でお知らせします。既定は **OFF** です（詳しくは「更新の確認」を参照）。
* **設定の自動保存**: 変更操作が落ち着いてから約500ms後、および設定画面を閉じる際に保存します。保存に失敗した場合は設定画面に表示します。
* **IME取得失敗への対応**: 入力モードを取得できない場合は、英数入力と誤表示せずインジケーターを隠します。

## 動作環境

* **OS**: Windows 11
* **ランタイム**: .NET 8.0 Desktop Runtime

## インストールと実行

### バイナリを使用する場合
リリースページからzipファイルをダウンロードして展開し、`Imel.exe` を実行してください。

### ソースコードからビルドする場合
1.  このリポジトリをクローンします。
2.  Visual Studio 2022で `Imel.csproj` を開きます。
3.  ソリューション構成を `Release` に設定し、ビルドします。
4.  `bin/Release/net8.0-windows` フォルダ内の `Imel.exe` を実行します。

Releaseビルドでは、同じフォルダに `README.md` もコピーされます。

### 設定の保存と更新

設定は `%APPDATA%\Imel\settings.json` に保存されます。既存の設定ファイルもそのまま読み込めます。「画面端で配置を反転」が未設定の場合はOFFとなり、「初期設定に戻す」でもOFFに戻ります。

保存時は一時ファイルへ書き込んでから置換し、正常な既存設定を `settings.json.bak` に退避します。本体が破損・欠落している場合はバックアップから読み込み、どちらも読み込めない場合は初期設定を使用します。

### 更新の確認

設定画面の「全般」にある「更新の確認」をONにすると、ONにした直後、アプリの起動時、24時間ごとに、GitHubで最新のリリースを確認します。前回の確認から24時間たっていない場合、起動時には確認しません。確認に失敗した場合は、1時間後にもう一度確認します。

* 新しいバージョンがあれば、Windowsの通知でお知らせします。同じバージョンの通知は1回だけです。通知をクリックすると、ブラウザーでリリースページを開きます。
* 更新があるあいだは、タスクトレイのメニューの一番上に「vX.Y.Z をダウンロード...」が表示されます。
* 設定画面では、最後に確認した日時と結果を確認でき、「今すぐ確認」で手動で確認できます。
* 更新のダウンロードやインストールは自動では行いません。リリースページからzipファイルをダウンロードして、置き換えてください。
* 確認のときは `api.github.com` に接続します。送信されるのは、通常のHTTPS通信の情報（IPアドレスなど）と、Imelのバージョンを含むUser-Agentだけです。OFFのあいだは通信しません。

## 更新履歴

### v1.0.15 (2026-10-06)

* **設定画面の改善**: スクリーンリーダーなどの支援技術で、スイッチ・スライダー・色の選択・数値の入力欄の名前を読み上げられるようにしました。
* **内部の整理**: 動作は変わりません。
    * Win32 APIの宣言とタスクトレイの処理を、`MainWindow` から別のファイルに分けました。
    * 処理に失敗した場合の診断出力を追加しました（デバッガーを接続したときに確認できます）。
* **開発環境**: `.editorconfig` を追加しました。`Imel.csproj` のバージョンとREADMEの表記が一致しているかを確かめるテストを追加し、回帰テストは81件になりました。

### v1.0.14 (2026-10-06)

* **表示文言の修正**:
    * 更新の通知とトレイのメニューで、「Imel」「vX.Y.Z」などと日本語の間にあった余分な空白を取りました。
    * 設定画面の「自動非表示」の説明を、「マウスカーソルが非表示のときに隠す（動画視聴中など）」に改めました。

### v1.0.13 (2026-10-06)

* **更新確認の改善**:
    * 前回の確認結果（新しいバージョン、失敗したかどうか）を保存し、再起動後もトレイのメニューと設定画面に引き継ぐようにしました。確認に失敗した場合は、再起動しても1時間後に再確認します。
    * 最後の確認日時が未来になっている場合（PCの時計の変更など）も、確認が止まらないようにしました。
* **フォーカスの修正**: インジケーターを表示するときに、入力中のアプリからフォーカスを奪わないようにしました。
* **設定画面の改善**:
    * 不透明度・オフセット・RGBの入力欄を空にした場合や、数値として読めない値を入力した場合は、直前の値を維持するようにしました（従来は100や0に変わっていました）。
    * 数値の入力欄と設定画面を広げ、入力した値が読めるようにしました。
* **ライブラリの更新**: WPF-UI を 4.1.0 から 4.3.0 に更新しました。
* **開発環境**: GitHub Actions によるビルドと回帰テストの自動実行、Dependabot による依存関係の更新確認を追加しました。回帰テストは78件になりました。

### v1.0.12 (2026-10-05)

* **更新の確認先を変更**: GitHubのアカウント名の変更に合わせて、更新を確認するリポジトリを `tabunugoku/Imel` に変更しました。v1.0.10・v1.0.11をお使いの場合は、v1.0.12へ更新してください。
* **著作権表記の変更**: LICENSEとアプリのバージョン情報の著作権者を `tabunugoku` に変更しました。
* **ビルドの変更**: 配布ファイルに、ビルドしたPCのフォルダのパスが含まれないようにしました。

### v1.0.11 (2026-10-05)

* **更新の確認**: 前回の確認から24時間（失敗した場合は1時間）たっていないときは、起動時に確認しないようにしました。
* **設定画面の改善**: 表示位置のオフセットと不透明度に範囲外の値を入力した場合、実際に適用される値を入力欄にも反映するようにしました。
* **細かな修正**: 終了時にタスクトレイのメニューを破棄するようにしました。

### v1.0.10 (2026-10-03)

* **更新の通知を追加**: 設定画面に「更新の確認」を追加しました。ONにすると、GitHubで新しいバージョンを確認し、Windowsの通知とタスクトレイのメニューでお知らせします。既定はOFFです。
* **回帰テスト**: バージョンの比較、リリース情報の解析、更新の確認に関する設定の保存と既定値を検証するテストを追加しました。

### v1.0.9 (2026-10-03)

* **設定画面の改善**:
    * 設定を保存できない場合も、確認のうえで設定画面を閉じられるようにしました。保存できなかった変更は、アプリの終了時にもう一度保存を試みます。
    * 保存エラーの文字色を、設定画面のテーマ（ライト/ダーク）に合わせて表示するようにしました。
    * RGBの入力値を0～255に収めてから反映し、範囲外の値で色が変わる問題を修正しました。
    * 位置オフセットの入力範囲を-200～200（DIP）に制限しました。既存の設定ファイルにある範囲外の値は、読み込み時に範囲内へ収めます。
* **終了時の警告の重複を修正**: 設定を保存できない状態でアプリを終了したとき、警告が2回表示される場合がある問題を修正しました。
* **スタートアップ登録の改善**:
    * 登録済みの `Imel.exe` が見つからない場合は、起動したときに現在の `Imel.exe` で登録し直すようにしました。
    * 登録先のレジストリキーがない場合も登録できるようにし、登録に失敗したときはスイッチを実際の登録状態に戻すようにしました。
* **開発環境**: 回帰テストのプロジェクトをソリューションに追加し、スタートアップ登録のパス解析とオフセットの範囲を検証するテストを追加しました。

### v1.0.8 (2026-09-26)

* **IME表示の修正**:
    * IME状態・変換モードの取得に失敗した場合、半角英数表示ではなく非表示に変更しました。
    * 非同期取得後に前面ウィンドウと入力先を再確認し、対象が変わっている場合は結果を破棄します。
    * 再表示前に位置を更新し、終了後の非同期処理による表示更新を抑止します。
* **設定保存の改善**:
    * 変更後の遅延保存と設定画面を閉じる際の保存を追加しました。
    * 一時ファイルからの置換保存、バックアップからの復旧、保存失敗の表示に対応しました。
* **非表示中の処理削減**:
    * 位置追従とIME状態確認のタイマーを分離しました。
    * 非表示中は位置追従タイマーを停止し、100ms間隔の状態確認を継続します。
* **画面・DPI対応**:
    * 物理座標とウィンドウのDPIを使う配置処理に変更し、Per-Monitor V2を明示しました。
    * 「画面端で配置を反転」の切り替えを追加しました。既定はOFFです。
    * 設定画面をリサイズ可能にし、初期表示時・DPI変更時に作業領域へ収める処理を追加しました。
    * 「透過率」の表記を「背景の不透明度」に修正しました。
* **回帰テスト**: IME取得、設定保存・復旧、配置計算、配置反転設定の保存と既定値を検証するテストを追加しました。

### v1.0.7 (2026-05-27)
* **【修正】設定画面を開いた後にインジケーターの見た目が変わる問題を修正**:
    * 設定画面のライト/ダークテーマ適用がアプリ全体へ波及し、メインインジケーターの背景色や角の描画に影響する問題を修正しました。
    * 設定画面だけにテーマリソースを適用するよう変更し、メインインジケーターの表示を設定画面から分離しました。
    * 設定画面を開いている間もインジケーターの追従が遅れにくいよう、位置更新タイマーの優先度を調整しました。
* **【修正】多重起動防止**:
    * 名前付きMutexを使用し、Imelが複数起動しないようにしました。

### v1.0.6 (2026-05-27)
* **軽量化と常駐時の負荷低減**:
    * IME状態・前面ウィンドウ・カーソル表示状態の確認頻度を見直し、常時実行される処理量を削減しました。
    * IME状態の問い合わせ処理をバックグラウンド化し、マウスカーソルへの追従処理が止まりにくいよう改善しました。
    * マウス位置が変化していない場合は表示位置を再設定しないようにし、WPF側の不要な更新を抑制しました。
    * 設定画面用のウィンドウアイコン画像の読み込みを遅延化し、通常の常駐時に読み込むリソースを削減しました。
    * 定期的な強制GCおよびワーキングセット削減処理を廃止し、.NETランタイム本来のメモリ管理に任せる設計へ変更しました。
* **設定の調整**:
    * 描画間隔の既定値を10msから16msに変更し、追従性を保ちながら省負荷寄りの初期設定にしました。
    * 設定画面の描画間隔に関する推奨値表記を更新しました。
    * 設定画面の配色がWindowsのアプリテーマ（ライト/ダーク）に追従するようにしました。

### v1.0.5 (2025-12-20)
* **【修正】** テーマ変更機能（ダークモード等）により表示不具合が発生していたため、機能を削除しv1.0.3相当の仕様にロールバックしました。
* 設定画面のレイアウトを微調整しました。

### v1.0.4 (2025-12-20)
* **設定画面の機能強化とデザイン刷新**:
    * **テーマ設定の実装**: アプリケーションの配色を「ライト」「ダーク」「システム設定」から選択可能になりました。
    * **Micaマテリアルの適用**: 設定ウィンドウの背景にMica効果を適用し、Windows 11のデスクトップになじむ見た目にしました。
    * **レイアウトの最適化**: 画面サイズをコンパクトに固定し、スクロール不要で全設定項目にアクセスできるようUIレイアウトを再調整しました。

### v1.0.3 (2025-12-18)
* **【重要】キーボード入力の安定性を向上**:
    * 特定の条件下で、キーボード入力が一時的に受け付けられなくなる（フリーズする）致命的な不具合を修正しました。
    * 内部ロジックを刷新し、`AttachThreadInput`（スレッド入力の共有）を廃止しました。代わりに `SendMessageTimeout` と `GetGUIThreadInfo` を組み合わせた非同期通信方式を採用しました。これにより、監視対象のアプリケーションが応答しない場合でも、Imelとシステム全体の入力処理が巻き込まれて止まらないように改善しました。

### v1.0.2 (2025-12-17)
* **設定画面UIのモダナイズ**: 設定画面のデザインとレイアウトを刷新し、モダンで洗練されたインターフェースにしました。
    * **一覧性の向上**: ウィンドウサイズを最適化し、スクロールなしですべての設定項目へアクセス可能になりました。
    * **操作性の改善**: RGBカラー設定や各項目の配置を見直し、より直感的にカスタマイズできるよう調整しました。
* **その他**: プロジェクト構成の最適化と、バージョン情報の更新を行いました。

## 開発環境

* Visual Studio 2022
* C# (WPF / .NET 8)

## 開発・テスト

自動テストは `dotnet run --project Tests/Imel.RegressionTests.csproj` で実行できます。テストのプロジェクトは `Imel.sln` にも含まれているため、ソリューションのビルドで一緒にビルドされます。
変更時には、ビルドとロジックの自動テストに加え、以下の実機動作も確認してください。

* 起動・終了、多重起動防止、タスクトレイ操作。
* アプリ切り替えと各IME入力モードの表示、カーソル非表示からの復帰。
* 設定変更・画面を閉じる・アプリ再起動後の設定復帰。
* 配置反転のON/OFF、既定OFF、初期設定へのリセット。
* 100%／150%／200%の混在DPI、負座標のモニター、画面端での追従と設定画面の表示。

非表示中のタイマー停止は実装済みですが、CPU使用率・消費電力の削減幅は未計測です。

## ライセンス

本ソフトウェアは [MIT License](LICENSE) の下で公開されています。
詳細は LICENSE ファイルをご確認ください。

</details>
