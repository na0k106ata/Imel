using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Imel
{
    internal sealed record UpdateInfo(Version Version, string ReleaseUrl);

    internal enum UpdateCheckState { NotChecked, Checking, UpToDate, UpdateAvailable, Failed }

    /// <summary>
    /// GitHub の最新リリースを問い合わせ、新しいバージョンがあるかを確認します。
    /// </summary>
    internal static class UpdateChecker
    {
        private const string LatestReleaseApi = "https://api.github.com/repos/tabunugoku/Imel/releases/latest";

        // 通知から開く URL は、このリポジトリのリリースページに限る。
        private const string ReleasePagePrefix = "https://github.com/tabunugoku/Imel/releases/";

        private static readonly Lazy<HttpClient> Http = new(CreateClient);

        private static HttpClient CreateClient()
        {
            var client = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
            client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("Imel", GetCurrentVersion().ToString()));
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
            return client;
        }

        internal static Version GetCurrentVersion()
        {
            var version = typeof(UpdateChecker).Assembly.GetName().Version ?? new Version(0, 0, 0);
            return new Version(version.Major, version.Minor, Math.Max(0, version.Build));
        }

        // 新しい版があれば UpdateInfo、最新なら null を返す。通信や解析に失敗した場合は例外を投げる。
        internal static async Task<UpdateInfo?> CheckAsync(CancellationToken cancellationToken)
        {
            using var response = await Http.Value.GetAsync(LatestReleaseApi, cancellationToken).ConfigureAwait(false);
            if (response.StatusCode == HttpStatusCode.NotFound) return null; // 公開済みのリリースがない
            response.EnsureSuccessStatusCode();

            string json = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            return ParseLatestRelease(json, GetCurrentVersion());
        }

        internal static UpdateInfo? ParseLatestRelease(string json, Version current)
        {
            using var document = JsonDocument.Parse(json);
            var root = document.RootElement;
            if (IsTrue(root, "draft") || IsTrue(root, "prerelease")) return null;

            if (!root.TryGetProperty("tag_name", out var tag) || tag.ValueKind != JsonValueKind.String ||
                !TryParseVersion(tag.GetString(), out var latest))
                throw new FormatException("リリースのバージョンを読み取れません。");

            if (latest <= current) return null;

            string? url = root.TryGetProperty("html_url", out var htmlUrl) && htmlUrl.ValueKind == JsonValueKind.String
                ? htmlUrl.GetString() : null;
            if (url == null || !url.StartsWith(ReleasePagePrefix, StringComparison.Ordinal))
                url = ReleasePagePrefix + "latest";

            return new UpdateInfo(latest, url);
        }

        // "v1.0.10" や "1.0.10" を 3 桁のバージョンとして読み取る。"v1.0.10-beta" などは対象外。
        internal static bool TryParseVersion(string? tag, out Version version)
        {
            version = new Version(0, 0, 0);
            if (string.IsNullOrWhiteSpace(tag)) return false;

            string text = tag.Trim();
            if (text.StartsWith('v') || text.StartsWith('V')) text = text[1..];
            if (!Version.TryParse(text, out var parsed) || parsed.Build < 0) return false;

            version = new Version(parsed.Major, parsed.Minor, parsed.Build);
            return true;
        }

        private static bool IsTrue(JsonElement element, string name) =>
            element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.True;
    }
}
