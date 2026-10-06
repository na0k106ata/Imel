using System;
using Imel;

int passed = 0;
var failures = new System.Collections.Generic.List<string>();
// 失敗しても最後まで実行し、全件の失敗を最後にまとめて表示する。
void Fail(string name, string detail)
{
    failures.Add(name);
    Console.WriteLine($"FAIL {name}: {detail}");
}
void Check(string name, bool openSuccess, bool isOpen, bool conversionSuccess,
    int mode, string? expected, int expectedCalls, bool hasIme = true, bool hasTarget = true)
{
    int calls = 0;
    bool Query(IntPtr window, int command, out IntPtr result)
    {
        calls++;
        if (window != (IntPtr)20) throw new Exception("Wrong IME window");
        if (command == 5) { result = isOpen ? (IntPtr)1 : IntPtr.Zero; return openSuccess; }
        if (command != 1) throw new Exception("Wrong command");
        result = (IntPtr)mode;
        return conversionSuccess;
    }
    string? actual;
    try
    {
        actual = ImeStatusReader.Read(hasTarget ? (IntPtr)10 : IntPtr.Zero,
            target => hasIme ? (IntPtr)20 : IntPtr.Zero, Query);
    }
    catch (Exception ex)
    {
        Fail(name, ex.Message);
        return;
    }
    if (actual != expected || calls != expectedCalls)
    {
        Fail(name, $"actual={actual ?? "null"}, calls={calls}");
        return;
    }
    Console.WriteLine($"PASS {name}");
    passed++;
}

Check("No foreground target", true, true, true, 9, null, 0, hasTarget: false);
Check("No IME window", true, true, true, 9, null, 0, hasIme: false);
Check("Open query timeout/failure", false, false, true, 9, null, 1);
Check("IME OFF skips conversion query", true, false, false, 9, "_A", 1);
Check("Conversion query timeout/failure is unknown", true, true, false, 0, null, 2);
Check("Hiragana", true, true, true, 9, "あ", 2);
Check("Full width katakana", true, true, true, 11, "カ", 2);
Check("Half width katakana", true, true, true, 3, "_ｶ", 2);
Check("Full width alphanumeric", true, true, true, 8, "Ａ", 2);
Check("Half width alphanumeric", true, true, true, 0, "_A", 2);
Check("Additional conversion flags", true, true, true, 9 | 0x10, "あ", 2);
Console.WriteLine($"{passed} checks passed");
void Assert(string name, bool condition)
{
    if (!condition)
    {
        Fail(name, "condition is false");
        return;
    }
    Console.WriteLine("PASS " + name);
    passed++;
}
var workArea = new PixelRect(0, 0, 1920, 1040);
var normal = IndicatorPlacement.Calculate(100, 100, 24, 24, 15, 15, 1, workArea);
Assert("Normal placement", normal == new PixelRect(115, 115, 139, 139));
var edge = IndicatorPlacement.Calculate(1919, 1039, 24, 24, 15, 15, 1, workArea, true);
Assert("Right and bottom flip", edge == new PixelRect(1880, 1000, 1904, 1024));
var negative = IndicatorPlacement.Calculate(-1, -1, 24, 24, 15, 15, 1.5, new PixelRect(-1920, -1080, 0, 0), true);
Assert("Negative monitor", negative.Right <= 0 && negative.Bottom <= 0 && negative.Width == 36);
foreach (double scale in new[] { 1.0, 1.5, 2.0 })
{
    var rect = IndicatorPlacement.Calculate(2100, 500, 24, 24, 15, 15, scale, new PixelRect(1920, 0, 3840, 2160));
    Assert("DPI " + scale, rect.Left == (int)Math.Round(2100 + 15 * scale) && rect.Width == (int)(24 * scale));
}
var extreme = IndicatorPlacement.Calculate(100, 100, 24, 24, int.MaxValue + 5.0, int.MinValue + 5.0, 2, workArea, true);
Assert("Extreme offsets are clamped", extreme.Left >= 0 && extreme.Right <= 1920 && extreme.Top >= 0 && extreme.Bottom <= 1040);
var tiny = IndicatorPlacement.Calculate(1, 1, 48, 48, 15, 15, 2, new PixelRect(0, 0, 10, 10), true);
Assert("Tiny screen", tiny == new PixelRect(0, 0, 10, 10));
var noFlip = IndicatorPlacement.Calculate(1919, 1039, 24, 24, 15, 15, 1, workArea);
Assert("Default OFF preserves offset past screen edge", noFlip == new PixelRect(1934, 1054, 1958, 1078));
Assert("Explicit OFF matches default", noFlip == IndicatorPlacement.Calculate(1919, 1039, 24, 24, 15, 15, 1, workArea, false));
var noFlipDpi = IndicatorPlacement.Calculate(-1, -1, 24, 24, 15, 15, 2, new PixelRect(-1920, -1080, 0, 0), false);
Assert("OFF preserves high DPI offset across boundary", noFlipDpi == new PixelRect(29, 29, 77, 77));
Assert("New settings default OFF", !new AppSettings().FlipAtScreenEdge);
Assert("Startup path quoted", StartupRegistration.GetExecutablePath(@"""C:\Apps\Imel.exe""") == @"C:\Apps\Imel.exe");
Assert("Startup path with arguments", StartupRegistration.GetExecutablePath(@"""C:\My Apps\Imel.exe"" --tray") == @"C:\My Apps\Imel.exe");
var current = new Version(1, 0, 9);
Assert("Version tag with v", UpdateChecker.TryParseVersion("v1.0.10", out var tagVersion) && tagVersion == new Version(1, 0, 10));
Assert("Version tag without v", UpdateChecker.TryParseVersion("1.2.3", out var plainVersion) && plainVersion == new Version(1, 2, 3));
Assert("Two part version tag rejected", !UpdateChecker.TryParseVersion("v1.0", out _));
Assert("Pre-release version tag rejected", !UpdateChecker.TryParseVersion("v1.0.10-beta", out _));
var newer = UpdateChecker.ParseLatestRelease(@"{""tag_name"":""v1.0.10"",""html_url"":""https://github.com/tabunugoku/Imel/releases/tag/v1.0.10""}", current);
Assert("Newer release found", newer != null && newer.Version == new Version(1, 0, 10) && newer.ReleaseUrl.EndsWith("/v1.0.10"));
Assert("Numeric comparison (1.0.10 > 1.0.9)", newer != null);
Assert("Same version is not an update", UpdateChecker.ParseLatestRelease(@"{""tag_name"":""v1.0.9""}", current) == null);
Assert("Older version is not an update", UpdateChecker.ParseLatestRelease(@"{""tag_name"":""v1.0.8""}", current) == null);
Assert("Draft is ignored", UpdateChecker.ParseLatestRelease(@"{""tag_name"":""v2.0.0"",""draft"":true}", current) == null);
Assert("Prerelease is ignored", UpdateChecker.ParseLatestRelease(@"{""tag_name"":""v2.0.0"",""prerelease"":true}", current) == null);
var foreign = UpdateChecker.ParseLatestRelease(@"{""tag_name"":""v2.0.0"",""html_url"":""https://example.com/evil""}", current);
Assert("Foreign release URL replaced", foreign != null && foreign.ReleaseUrl == "https://github.com/tabunugoku/Imel/releases/latest");
bool invalidTagThrows = false;
try { UpdateChecker.ParseLatestRelease(@"{""tag_name"":""latest""}", current); }
catch (FormatException) { invalidTagThrows = true; }
Assert("Invalid tag reports failure", invalidTagThrows);
Assert("New settings do not check for updates", !new AppSettings().CheckForUpdates);
var now = new DateTime(2026, 10, 6, 12, 0, 0, DateTimeKind.Utc);
Assert("Never checked is due", UpdateChecker.IsCheckDue(null, false, now));
Assert("Recent check is not due", !UpdateChecker.IsCheckDue(now.AddHours(-23), false, now));
Assert("24 hours later is due", UpdateChecker.IsCheckDue(now.AddHours(-24), false, now));
Assert("Recent failure waits for retry", !UpdateChecker.IsCheckDue(now.AddMinutes(-30), true, now));
Assert("Failure retries after 1 hour", UpdateChecker.IsCheckDue(now.AddHours(-1), true, now));
Assert("Failure retry survives restart (not 24h)", UpdateChecker.IsCheckDue(now.AddHours(-2), true, now));
Assert("Future check time is due", UpdateChecker.IsCheckDue(now.AddDays(30), false, now));
var restored = UpdateChecker.RestoreUpdate("1.0.10", "https://github.com/tabunugoku/Imel/releases/tag/v1.0.10", current);
Assert("Saved update restored", restored != null && restored.Version == new Version(1, 0, 10) && restored.ReleaseUrl.EndsWith("/v1.0.10"));
Assert("Saved update already installed is dropped", UpdateChecker.RestoreUpdate("1.0.9", null, current) == null);
Assert("Saved update missing is dropped", UpdateChecker.RestoreUpdate(null, null, current) == null);
Assert("Saved update garbage is dropped", UpdateChecker.RestoreUpdate("abc", null, current) == null);
Assert("Saved foreign URL replaced", UpdateChecker.RestoreUpdate("2.0.0", "https://example.com/x", current)?.ReleaseUrl == "https://github.com/tabunugoku/Imel/releases/latest");
Assert("Startup path unquoted", StartupRegistration.GetExecutablePath(@"  C:\Apps\Imel.exe ") == @"C:\Apps\Imel.exe");

string directory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ImelTests-" + Guid.NewGuid().ToString("N"));
System.IO.Directory.CreateDirectory(directory);
string path = System.IO.Path.Combine(directory, "settings.json");
try
{
    Assert("Missing settings use defaults", AppSettings.Load(path).OffsetX == 10);
    Assert("First save", AppSettings.Save(new AppSettings { OffsetX = 12 }, out _, path));
    Assert("First save roundtrip", AppSettings.Load(path).OffsetX == 12);
    Assert("Replace save", AppSettings.Save(new AppSettings { OffsetX = 34 }, out _, path));
    Assert("Latest data", AppSettings.Load(path).OffsetX == 34);
    Assert("Backup contains previous data", AppSettings.Load(path + ".bak").OffsetX == 12);
    System.IO.File.WriteAllText(path, "{corrupt");
    Assert("Corrupt primary recovers backup", AppSettings.Load(path).OffsetX == 12);
    Assert("Save after recovery", AppSettings.Save(new AppSettings { OffsetX = 56 }, out _, path));
    Assert("Recovery preserves good backup", AppSettings.Load(path + ".bak").OffsetX == 12);
    using (var locked = new System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None))
    {
        Assert("Locked destination reports failure", !AppSettings.Save(new AppSettings { OffsetX = 99 }, out string? error, path) && !string.IsNullOrEmpty(error));
    }
    Assert("Failed save preserves original", AppSettings.Load(path).OffsetX == 56);
    Assert("Temporary files cleaned", System.IO.Directory.GetFiles(directory, "*.tmp").Length == 0);
    System.IO.File.WriteAllText(path, "{\"Scale\":20,\"Opacity\":-1,\"UpdateInterval\":0}");
    var normalized = AppSettings.Load(path);
    Assert("Invalid values normalized", normalized.Scale == 2 && normalized.Opacity == 0 && normalized.UpdateInterval == 2);
    System.IO.File.WriteAllText(path, "{\"OffsetX\":100000,\"OffsetY\":-100000}");
    var offsets = AppSettings.Load(path);
    Assert("Out of range offsets normalized", offsets.OffsetX == AppSettings.MaxOffset && offsets.OffsetY == AppSettings.MinOffset);
    System.IO.File.WriteAllText(path, "null");
    Assert("Null JSON recovers backup", AppSettings.Load(path).OffsetX == 12);
    System.IO.File.WriteAllText(path + ".bak", "{corrupt");
    Assert("Both corrupt use defaults", AppSettings.Load(path).OffsetX == 10);
    System.IO.File.WriteAllText(path, "{\"OffsetX\":42}");
    Assert("Existing settings without new property default OFF", !AppSettings.Load(path).FlipAtScreenEdge && AppSettings.Load(path).OffsetX == 42);
    Assert("Save ON", AppSettings.Save(new AppSettings { FlipAtScreenEdge = true }, out _, path));
    Assert("ON survives reload", AppSettings.Load(path).FlipAtScreenEdge);
    Assert("Save OFF", AppSettings.Save(new AppSettings { FlipAtScreenEdge = false }, out _, path));
    Assert("OFF survives reload", !AppSettings.Load(path).FlipAtScreenEdge);
    var checkedAt = new DateTime(2026, 10, 3, 9, 30, 0, DateTimeKind.Utc);
    Assert("Save update settings", AppSettings.Save(new AppSettings { CheckForUpdates = true, LastUpdateCheckUtc = checkedAt, NotifiedUpdateVersion = "1.0.10" }, out _, path));
    var updateSettings = AppSettings.Load(path);
    Assert("Update settings survive reload", updateSettings.CheckForUpdates && updateSettings.LastUpdateCheckUtc == checkedAt && updateSettings.NotifiedUpdateVersion == "1.0.10");
    Assert("Save update result", AppSettings.Save(new AppSettings { CheckForUpdates = true, LastUpdateCheckFailed = true, AvailableUpdateVersion = "1.0.13", AvailableUpdateUrl = "https://github.com/tabunugoku/Imel/releases/tag/v1.0.13" }, out _, path));
    var updateResult = AppSettings.Load(path);
    Assert("Update result survives reload", updateResult.LastUpdateCheckFailed && updateResult.AvailableUpdateVersion == "1.0.13" && updateResult.AvailableUpdateUrl?.EndsWith("v1.0.13") == true);
    System.IO.File.WriteAllText(path, "{\"OffsetX\":42}");
    Assert("Existing settings without update option default OFF", !AppSettings.Load(path).CheckForUpdates);
    Assert("Existing settings without update result have none", !AppSettings.Load(path).LastUpdateCheckFailed && AppSettings.Load(path).AvailableUpdateVersion == null);
}
catch (Exception ex)
{
    Fail("Unexpected exception in settings tests", ex.ToString());
}
finally
{
    // このテストが作成した直下のファイルのみ削除する。
    foreach (string file in System.IO.Directory.GetFiles(directory))
        System.IO.File.Delete(file);
    System.IO.Directory.Delete(directory);
}
if (failures.Count > 0)
{
    Console.WriteLine($"{failures.Count} FAILED, {passed} passed");
    foreach (string name in failures) Console.WriteLine("  - " + name);
    return 1;
}
Console.WriteLine($"{passed} total checks passed");
return 0;
