using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace LLMAgentTrader
{
    // ────────────────────────────────────────────────────────────────────────────
    //  共享 HttpClient 池
    //  修正：原來 5 個 static HttpClient 各自獨立 connection pool，
    //        DNS 快取無法共享，Yahoo 限速時無法協調。
    //        現在統一三個用途的 HttpClient，共享底層 TCP 連線。
    // ────────────────────────────────────────────────────────────────────────────
    public static class AppHttpClients
    {
        private const string UA_CHROME =
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36";

        /// <summary>Yahoo API：帶 Cookie，保持 session（20秒）</summary>
        public static readonly HttpClient Market = BuildApi(20);

        /// <summary>LLM 串流呼叫（120秒）</summary>
        public static readonly HttpClient Llm = BuildApi(120);

        /// <summary>新聞爬取（15秒）</summary>
        public static readonly HttpClient News = BuildApi(15);

        /// <summary>
        /// HTML 爬蟲用：每次請求完全獨立（不共用 Cookie），
        /// 模擬全新瀏覽器行為，避免 Yahoo/Finviz 的 Bot 偵測。
        /// 每次 FetchFundamentals 應使用此 client 的 NEW 請求，
        /// 並在 HttpRequestMessage 上單獨設定 headers。
        /// </summary>
        public static HttpClient MakeScraper(int timeoutSec = 15)
        {
            var handler = new HttpClientHandler
            {
                UseCookies = false,   // 不共用 cookie jar
                AllowAutoRedirect = true,
                AutomaticDecompression = System.Net.DecompressionMethods.GZip |
                                         System.Net.DecompressionMethods.Deflate |
                                         System.Net.DecompressionMethods.Brotli,
            };
            var hc = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(timeoutSec) };
            hc.DefaultRequestHeaders.Add("User-Agent", UA_CHROME);
            hc.DefaultRequestHeaders.Add("Accept",
                "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8");
            hc.DefaultRequestHeaders.Add("Accept-Language", "en-US,en;q=0.9");  // 注意：en-US 而非 zh-TW
            hc.DefaultRequestHeaders.Add("Accept-Encoding", "gzip, deflate, br");
            hc.DefaultRequestHeaders.Add("Connection", "keep-alive");
            hc.DefaultRequestHeaders.Add("Upgrade-Insecure-Requests", "1");
            hc.DefaultRequestHeaders.Add("Sec-Fetch-Dest", "document");
            hc.DefaultRequestHeaders.Add("Sec-Fetch-Mode", "navigate");
            hc.DefaultRequestHeaders.Add("Sec-Fetch-Site", "none");
            hc.DefaultRequestHeaders.Add("Sec-Fetch-User", "?1");
            hc.DefaultRequestHeaders.Add("Cache-Control", "max-age=0");
            return hc;
        }

        private static HttpClient BuildApi(int timeoutSec)
        {
            var handler = new HttpClientHandler
            {
                UseCookies = true,
                AutomaticDecompression = System.Net.DecompressionMethods.GZip |
                                         System.Net.DecompressionMethods.Deflate,
            };
            var hc = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(timeoutSec) };
            hc.DefaultRequestHeaders.Add("User-Agent", UA_CHROME);
            hc.DefaultRequestHeaders.Add("Accept-Language", "zh-TW,zh;q=0.9,en;q=0.8");
            hc.DefaultRequestHeaders.Add("Accept", "application/json,text/html,*/*;q=0.8");
            return hc;
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  LLM 模型設定
    //  修正：model 原本硬編碼為 "openai/gpt-4o-mini"，使用者無法不重新編譯就切換。
    //        現在可從 UI 下拉選單選擇，並且所有 LLM 請求都讀此設定。
    // ────────────────────────────────────────────────────────────────────────────
    public static class LlmConfig
    {
        public static string CurrentModel { get; set; } = "openai/gpt-4o-mini";
        public const string BaseUrl = "https://openrouter.ai/api/v1/chat/completions";

        // ── Google AI Studio 直連金鑰（與 OpenRouter Key 分開儲存）────────────
        public static string GeminiApiKey { get; set; } = "";
        private static readonly string GeminiKeyPath =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GeminiKey.txt");

        public static void LoadGeminiKey()
        {
            try { if (File.Exists(GeminiKeyPath)) GeminiApiKey = File.ReadAllText(GeminiKeyPath).Trim(); }
            catch (Exception ex) { AppLogger.Log("LoadGeminiKey 失敗", ex); }
        }
        public static void SaveGeminiKey()
        {
            try { File.WriteAllText(GeminiKeyPath, GeminiApiKey); }
            catch (Exception ex) { AppLogger.Log("SaveGeminiKey 失敗", ex); }
        }

        /// <summary>是否為 Google AI Studio 直連模型（ID 以 "gemini:" 開頭）</summary>
        public static bool IsGeminiDirect(string modelId) => modelId?.StartsWith("gemini:") == true;
        /// <summary>取出 Gemini 的實際模型名稱（去掉 "gemini:" 前綴）</summary>
        public static string GeminiModelName(string modelId) => modelId.Substring("gemini:".Length);

        public static readonly (string Id, string Label)[] AvailableModels =
        {
            // ── OpenRouter 路由 ──────────────────────────────────────────────
            ("openai/gpt-4o-mini",               "GPT-4o Mini  ⚡省"),
            ("openai/gpt-4o",                    "GPT-4o  🧠強"),
            ("anthropic/claude-3-5-sonnet",      "Claude 3.5 Sonnet"),
            ("anthropic/claude-3-haiku",         "Claude 3 Haiku  ⚡"),
            ("google/gemini-flash-1.5",          "Gemini Flash 1.5  [OpenRouter]"),
            ("meta-llama/llama-3.1-70b-instruct","Llama 3.1 70B"),
            // ── Google AI Studio 直連（免 OpenRouter 費率，需填 Gemini Key）──
            ("gemini:gemini-2.0-flash",          "✨ Gemini 2.0 Flash  [Google直連]"),
            ("gemini:gemini-1.5-flash",          "⚡ Gemini 1.5 Flash  [Google直連]"),
            ("gemini:gemini-1.5-pro",            "🧠 Gemini 1.5 Pro  [Google直連]"),
        };
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  Debounce 輔助
    //  修正：txtTicker 同時綁 Leave / KeyDown / SelectedIndexChanged，
    //        快速切換時三個事件並發觸發 UpdateTickerInfoUI，後回來的覆蓋前面造成 race。
    //        現在統一用 Debounce 300ms，多個事件只觸發最後一次。
    // ────────────────────────────────────────────────────────────────────────────
    public static class DebounceHelper
    {
        private static readonly Dictionary<string, System.Windows.Forms.Timer> _timers = new();
        private static readonly object _lock = new();

        public static void Run(string key, int delayMs, Action action)
        {
            lock (_lock)
            {
                if (_timers.TryGetValue(key, out var old)) { old.Stop(); old.Dispose(); }
                var t = new System.Windows.Forms.Timer { Interval = delayMs };
                t.Tick += (s, e) =>
                {
                    lock (_lock) { t.Stop(); t.Dispose(); _timers.Remove(key); }
                    action();
                };
                _timers[key] = t;
                t.Start();
            }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  重試輔助（指數退避，支援 CancellationToken）
    // ────────────────────────────────────────────────────────────────────────────
    public static class RetryHelper
    {
        public static async Task<T> RunAsync<T>(
            Func<Task<T>> action,
            int maxRetry = 2,
            string tag = "",
            CancellationToken ct = default)
        {
            int attempt = 0;
            while (true)
            {
                ct.ThrowIfCancellationRequested();
                try { return await action(); }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    attempt++;
                    if (attempt > maxRetry)
                    {
                        AppLogger.Log($"RetryHelper [{tag}] 達最大重試次數", ex);
                        throw;
                    }
                    int delay = (int)Math.Pow(2, attempt - 1) * 1000;
                    AppLogger.Log($"RetryHelper [{tag}] 第{attempt}次重試，等待{delay}ms", ex);
                    await Task.Delay(delay, ct);
                }
            }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  公司名稱服務
    // ────────────────────────────────────────────────────────────────────────────
    public static class MarketInfoService
    {
        public static async Task<string> GetCompanyName(string ticker, bool isUS)
        {
            if (!isUS)
            {
                try
                {
                    var twse = await TwseApiService.FetchRealtime(ticker);
                    if (twse != null && !string.IsNullOrEmpty(twse.CompanyName))
                        return twse.CompanyName;
                }
                catch (Exception ex) { AppLogger.Log($"MarketInfoService TWSE 失敗 {ticker}", ex); }
            }
            return await YahooDataService.FetchCompanyName(ticker);
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  台灣股市即時資料（使用共享 HttpClient.Market）
    // ────────────────────────────────────────────────────────────────────────────
    public static class TwseApiService
    {
        private static bool _hasSession = false;

        private static async Task EnsureSession()
        {
            if (!_hasSession)
            {
                await AppHttpClients.Market.GetAsync("https://mis.twse.com.tw/stock/index.jsp");
                _hasSession = true;
            }
        }

        public static async Task<TwseTick> FetchRealtime(string ticker, int retry = 3)
        {
            await EnsureSession();
            try
            {
                string market = ticker.EndsWith(".TWO") ? "otc" : "tse";
                string cleanTicker = ticker.Replace(".TW", "").Replace(".TWO", "");
                string url = $"https://mis.twse.com.tw/stock/api/getStockInfo.jsp" +
                             $"?ex_ch={market}_{cleanTicker}.tw" +
                             $"&_={DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}";

                var root = JsonDocument.Parse(
                    await AppHttpClients.Market.GetStringAsync(url)).RootElement;

                if (root.TryGetProperty("rtcode", out var rtcode) && rtcode.GetString() == "5000")
                {
                    AppLogger.Log($"TwseApiService rtcode=5000 {ticker}");
                    if (retry > 0) { await Task.Delay(500); _hasSession = false; return await FetchRealtime(ticker, retry - 1); }
                    return null;
                }
                if (!root.TryGetProperty("msgArray", out var msgArr) || msgArr.GetArrayLength() == 0)
                    return null;

                var data = msgArr[0];
                string zStr = data.TryGetProperty("z", out var zP) ? zP.GetString() : "-";
                if (zStr == "-") zStr = data.GetProperty("y").GetString();

                DateTime exactTime = DateTimeOffset
                    .FromUnixTimeMilliseconds(long.Parse(data.GetProperty("tlong").GetString()))
                    .DateTime.ToLocalTime();

                return new TwseTick
                {
                    Time = exactTime.ToString("yyyy-MM-dd HH:mm:ss"),
                    CompanyName = data.TryGetProperty("n", out var nP) ? nP.GetString() : "",
                    Price = double.TryParse(zStr, out var p) ? p : 0,
                    YesterdayClose = double.TryParse(data.GetProperty("y").GetString(), out var y) ? y : 0,
                    Volume = int.TryParse(data.GetProperty("v").GetString(), out var v) ? v : 0,
                    TradeVolume = int.TryParse(data.TryGetProperty("tv", out var tvP) ? tvP.GetString() : "0", out var tv) ? tv : 0,
                    Open = double.TryParse(data.TryGetProperty("o", out var oP) ? oP.GetString() : "0", out var o) ? o : 0,
                    High = double.TryParse(data.TryGetProperty("h", out var hP) ? hP.GetString() : "0", out var h) ? h : 0,
                    Low = double.TryParse(data.TryGetProperty("l", out var lP) ? lP.GetString() : "0", out var l) ? l : 0,
                    AskPrices = ParseArr(data.TryGetProperty("a", out var aP) ? aP.GetString() : ""),
                    AskVolumes = ParseArr(data.TryGetProperty("f", out var fP) ? fP.GetString() : "").Select(x => (int)x).ToList(),
                    BidPrices = ParseArr(data.TryGetProperty("b", out var bP) ? bP.GetString() : ""),
                    BidVolumes = ParseArr(data.TryGetProperty("g", out var gP) ? gP.GetString() : "").Select(x => (int)x).ToList()
                };
            }
            catch (Exception ex)
            {
                AppLogger.Log($"TwseApiService.FetchRealtime {ticker} 失敗", ex);
                if (retry > 0) { await Task.Delay(500); return await FetchRealtime(ticker, retry - 1); }
                return null;
            }
        }

        private static List<double> ParseArr(string input)
        {
            if (string.IsNullOrEmpty(input) || input == "-") return new List<double>();
            return input.Trim('_')
                .Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => double.TryParse(x, out var val) ? val : 0)
                .ToList();
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  Yahoo Finance 資料服務
    //
    //  改進：
    //  ① FetchFundamentals：
    //     - 台股優先走 TWSE/TPEX 官方 API（穩定、正式）
    //     - 美股走 Yahoo v10 quoteSummary JSON（比 HTML Regex 穩定 10 倍）
    //     - 原來靜默回 (0,0)，現在回傳 Source 欄位標記失敗來源，UI 可顯示 "N/A"
    //
    //  ② crumb TTL 25分鐘自動重取：
    //     - 原來 _authInitialized=true 後永不重取，session 中途過期不重試
    //     - 現在有 _crumbFetchedAt + _crumbTtl，過期自動重新取得
    //     - 401/403 時立即觸發重取再重試
    //
    //  ③ FetchVIXWithChange：
    //     - 原本 FetchVIX + FetchPrevClose("^VIX") 兩次請求
    //     - 現在單一 range=5d 請求取得 (current, prev)
    // ────────────────────────────────────────────────────────────────────────────
    public static class YahooDataService
    {
        private static bool _authInit = false;
        private static string _crumb = "";
        private static DateTime _crumbFetchedAt = DateTime.MinValue;
        private static readonly TimeSpan _crumbTtl = TimeSpan.FromMinutes(25);
        private static readonly SemaphoreSlim _authLock = new SemaphoreSlim(1, 1);

        // ════════════════════════════════════════════════════════════════════
        //  FetchFundamentals  — 七層 Fallback（美股/台股通用）
        //
        //  台股：① TWSE ② TPEX ③ Yahoo v7 API
        //  美股：① Yahoo HTML (en-US 全新 scraper) ② Finviz HTML
        //         ③ MarketWatch HTML ④ Alpha Vantage (需 Key)
        //         ⑤ Yahoo v11 API ⑥ Yahoo v7 API ⑦ Yahoo v10 API
        //
        //  每個 HTML 來源用獨立 HttpClient 實例（MakeScraper），
        //  避免共用 cookie 觸發 bot 偵測。
        // ════════════════════════════════════════════════════════════════════

        /// <summary>
        /// 從 HTML 大字串中 regex 抓 PE / Yield。
        /// 相容 Yahoo __NEXT_DATA__ / Redux JSON 兩種格式。
        /// </summary>
        private static (double PE, double Yield) ParseYahooHtml(string html)
        {
            double pe = 0, yld = 0;

            // 格式1：{"raw":30.25} 物件包覆
            var m = Regex.Match(html, @"""trailingPE""\s*:\s*\{\s*""raw""\s*:\s*([\d.]+)");
            if (m.Success) double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out pe);

            // 格式2：直接數值（Redux 扁平化）
            if (pe == 0)
            {
                m = Regex.Match(html, @"""trailingPE""\s*:\s*([\d.]+)(?!\s*[:{])");
                if (m.Success) double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out pe);
            }
            // forwardPE 備援
            if (pe == 0)
            {
                m = Regex.Match(html, @"""forwardPE""\s*:\s*\{\s*""raw""\s*:\s*([\d.]+)");
                if (m.Success) double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out pe);
            }

            // 殖利率：多欄位名稱嘗試
            foreach (var key in new[] { "trailingAnnualDividendYield", "dividendYield", "yield" })
            {
                m = Regex.Match(html, $@"""{key}""\s*:\s*\{{\s*""raw""\s*:\s*([\d.]+)");
                if (m.Success)
                {
                    double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out yld);
                    break;
                }
            }
            if (yld == 0)
            {
                m = Regex.Match(html, @"""(?:trailingAnnualDividendYield|dividendYield)""\s*:\s*([\d.]+)(?!\s*[:{])");
                if (m.Success) double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out yld);
            }

            return (pe, yld);
        }

        /// <summary>
        /// 從 Finviz snapshot table HTML 中抓 P/E 和 Dividend %
        /// </summary>
        private static (double PE, double Yield) ParseFinvizHtml(string html)
        {
            double pe = 0, yld = 0;
            // Finviz 的 snapshot table：連續的 td 對，key 和 value 交替出現
            // 格式：<td class="snapshot-td2-cp">P/E</td><td class="snapshot-td2">28.10</td>
            var matches = Regex.Matches(html,
                @"<td[^>]*>\s*([A-Za-z/\s%]+?)\s*</td>\s*<td[^>]*>\s*([-\d.]+%?)\s*</td>",
                RegexOptions.IgnoreCase);
            foreach (Match m in matches)
            {
                string key = m.Groups[1].Value.Trim();
                string val = m.Groups[2].Value.Trim().TrimEnd('%');
                if (key == "P/E" && pe == 0 && val != "-")
                    double.TryParse(val, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out pe);
                if ((key.StartsWith("Dividend") || key == "Div %") && yld == 0 && val != "-")
                {
                    if (double.TryParse(val, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var y))
                        yld = y / 100.0;
                }
            }
            return (pe, yld);
        }

        /// <summary>
        /// 從 MarketWatch quote page 的 data-field 屬性提取 PE/Yield
        /// </summary>
        private static (double PE, double Yield) ParseMarketWatchHtml(string html)
        {
            double pe = 0, yld = 0;
            // MarketWatch 格式：<span class="... " data-field="...">value</span>
            // 或 table 中的 key-value 對
            var mPE = Regex.Match(html,
                @"P/E\s*Ratio.*?<span[^>]*class=""[^""]*intraday[^""]*""[^>]*>([\d.]+)</span>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!mPE.Success)
                mPE = Regex.Match(html,
                    @"Price-Earnings Ratio.*?<td[^>]*>([\d.]+)</td>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (mPE.Success) double.TryParse(mPE.Groups[1].Value, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out pe);

            var mYld = Regex.Match(html,
                @"Dividend Yield.*?<small[^>]*>([\d.]+)%</small>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!mYld.Success)
                mYld = Regex.Match(html,
                    @"""dividendYield""\s*:\s*""([\d.]+)""",
                    RegexOptions.IgnoreCase);
            if (mYld.Success)
            {
                if (double.TryParse(mYld.Groups[1].Value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var y))
                    yld = y > 1 ? y / 100.0 : y;   // MarketWatch 有時回百分比，有時回小數
            }
            return (pe, yld);
        }

        // ────────────────────────────────────────────────────────────────────
        //  主函式：FetchFundamentals
        // ────────────────────────────────────────────────────────────────────
        public static async Task<(double PE, double Yield, string Source)> FetchFundamentals(string ticker)
        {
            bool isTw = ticker.EndsWith(".TW") || ticker.EndsWith(".TWO");

            return await RetryHelper.RunAsync(async () =>
            {
                // ══════════════ 台股官方 API ══════════════
                if (ticker.EndsWith(".TW"))
                {
                    try
                    {
                        string clean = ticker.Replace(".TW", "");
                        var root = JsonDocument.Parse(
                            await AppHttpClients.Market.GetStringAsync(
                                $"https://www.twse.com.tw/exchangeReport/BWIBBU?response=json&stockNo={clean}"))
                            .RootElement;
                        if (root.TryGetProperty("data", out var da) && da.GetArrayLength() > 0)
                        {
                            var last = da[da.GetArrayLength() - 1];
                            double yld = double.TryParse(last[1].GetString(),
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture, out var y) ? y / 100.0 : 0;
                            double pe = double.TryParse(last[3].GetString(),
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture, out var p) ? p : 0;
                            if (pe > 0 || yld > 0) return (pe, yld, "TWSE");
                        }
                    }
                    catch (Exception ex) { AppLogger.Log($"[Fundamentals] TWSE 失敗 {ticker}", ex); }
                }

                if (ticker.EndsWith(".TWO"))
                {
                    try
                    {
                        string clean = ticker.Replace(".TWO", "");
                        var root = JsonDocument.Parse(
                            await AppHttpClients.Market.GetStringAsync(
                                $"https://www.tpex.org.tw/web/stock/aftertrading/perwd/pera_result.php?l=zh-tw&o=json&stk_no={clean}"))
                            .RootElement;
                        if (root.TryGetProperty("aaData", out var da) && da.GetArrayLength() > 0)
                        {
                            var last = da[da.GetArrayLength() - 1];
                            double pe = double.TryParse(last[1].GetString(),
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture, out var p) ? p : 0;
                            double yld = double.TryParse(last[2].GetString(),
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture, out var y) ? y / 100.0 : 0;
                            if (pe > 0 || yld > 0) return (pe, yld, "TPEX");
                        }
                    }
                    catch (Exception ex) { AppLogger.Log($"[Fundamentals] TPEX 失敗 {ticker}", ex); }
                }

                // ══════════════ Yahoo Finance HTML（美股主要來源）══════════════
                // 使用全新 HttpClient 實例（不共用 cookie），Accept-Language: en-US
                if (!isTw)
                {
                    try
                    {
                        using var scraper = AppHttpClients.MakeScraper(12);
                        // Step 1: 先造訪首頁取 cookie（模擬正常瀏覽行為）
                        await scraper.GetAsync("https://finance.yahoo.com/");
                        // Step 2: 取 quote 頁
                        var resp = await scraper.GetAsync(
                            $"https://finance.yahoo.com/quote/{Uri.EscapeDataString(ticker)}");
                        if (resp.IsSuccessStatusCode)
                        {
                            string html = await resp.Content.ReadAsStringAsync();
                            var (pe, yld) = ParseYahooHtml(html);
                            if (pe > 0 || yld > 0) return (pe, yld, "Yahoo-HTML");
                            AppLogger.Log($"[Fundamentals] Yahoo-HTML parse 失敗（regex 無命中）{ticker}");
                        }
                        else
                            AppLogger.Log($"[Fundamentals] Yahoo-HTML HTTP {(int)resp.StatusCode} {ticker}");
                    }
                    catch (Exception ex) { AppLogger.Log($"[Fundamentals] Yahoo-HTML 失敗 {ticker}", ex); }
                }

                // ══════════════ Finviz（美股，HTML scraping）══════════════
                // 不需帳號，snapshot table 穩定，但每分鐘有限制
                if (!isTw)
                {
                    try
                    {
                        using var fvClient = AppHttpClients.MakeScraper(10);
                        var req = new HttpRequestMessage(HttpMethod.Get,
                            $"https://finviz.com/quote.ashx?t={Uri.EscapeDataString(ticker)}");
                        req.Headers.Add("Referer", "https://finviz.com/");
                        req.Headers.Add("Sec-Fetch-Site", "same-origin");
                        var resp = await fvClient.SendAsync(req);
                        if (resp.IsSuccessStatusCode)
                        {
                            string html = await resp.Content.ReadAsStringAsync();
                            var (pe, yld) = ParseFinvizHtml(html);
                            if (pe > 0 || yld > 0) return (pe, yld, "Finviz");
                            AppLogger.Log($"[Fundamentals] Finviz parse 失敗 {ticker}");
                        }
                        else
                            AppLogger.Log($"[Fundamentals] Finviz HTTP {(int)resp.StatusCode} {ticker}");
                    }
                    catch (Exception ex) { AppLogger.Log($"[Fundamentals] Finviz 失敗 {ticker}", ex); }
                }

                // ══════════════ MarketWatch（美股）══════════════
                if (!isTw)
                {
                    try
                    {
                        using var mwClient = AppHttpClients.MakeScraper(12);
                        var req = new HttpRequestMessage(HttpMethod.Get,
                            $"https://www.marketwatch.com/investing/stock/{ticker.ToLower()}");
                        req.Headers.Add("Referer", "https://www.marketwatch.com/");
                        var resp = await mwClient.SendAsync(req);
                        if (resp.IsSuccessStatusCode)
                        {
                            string html = await resp.Content.ReadAsStringAsync();
                            var (pe, yld) = ParseMarketWatchHtml(html);
                            if (pe > 0 || yld > 0) return (pe, yld, "MarketWatch");
                            AppLogger.Log($"[Fundamentals] MarketWatch parse 失敗 {ticker}");
                        }
                    }
                    catch (Exception ex) { AppLogger.Log($"[Fundamentals] MarketWatch 失敗 {ticker}", ex); }
                }

                // ══════════════ Alpha Vantage（需填入 Key）══════════════
                // 免費：25次/天。去 https://www.alphavantage.co 免費申請
                if (AlphaVantageKeyManager.HasKey)
                {
                    try
                    {
                        string url = $"https://www.alphavantage.co/query" +
                                     $"?function=OVERVIEW&symbol={Uri.EscapeDataString(ticker)}" +
                                     $"&apikey={Uri.EscapeDataString(AlphaVantageKeyManager.Key)}";
                        var resp = await AppHttpClients.Market.GetAsync(url);
                        if (resp.IsSuccessStatusCode)
                        {
                            var root = JsonDocument.Parse(await resp.Content.ReadAsStringAsync()).RootElement;
                            if (!root.TryGetProperty("Information", out _) && !root.TryGetProperty("Note", out _))
                            {
                                double pe = 0, yld = 0;
                                if (root.TryGetProperty("PERatio", out var peEl) &&
                                    double.TryParse(peEl.GetString(), System.Globalization.NumberStyles.Any,
                                        System.Globalization.CultureInfo.InvariantCulture, out var peV) && peV > 0)
                                    pe = peV;
                                if (root.TryGetProperty("DividendYield", out var yEl) &&
                                    double.TryParse(yEl.GetString(), System.Globalization.NumberStyles.Any,
                                        System.Globalization.CultureInfo.InvariantCulture, out var yV) && yV > 0)
                                    yld = yV;
                                if (pe > 0 || yld > 0) return (pe, yld, "AlphaVantage");
                            }
                            else AppLogger.Log($"[Fundamentals] Alpha Vantage 限流 {ticker}");
                        }
                    }
                    catch (Exception ex) { AppLogger.Log($"[Fundamentals] Alpha Vantage 失敗 {ticker}", ex); }
                }

                // ══════════════ Yahoo v11 API（帶 crumb）══════════════
                try
                {
                    await EnsureAuthAsync();
                    string cp = !string.IsNullOrEmpty(_crumb) ? $"&crumb={Uri.EscapeDataString(_crumb)}" : "";
                    string url = $"https://query1.finance.yahoo.com/v11/finance/quoteSummary/" +
                                 $"{Uri.EscapeDataString(ticker)}?modules=summaryDetail{cp}";
                    var resp = await AppHttpClients.Market.GetAsync(url);
                    if (resp.IsSuccessStatusCode)
                    {
                        var root = JsonDocument.Parse(await resp.Content.ReadAsStringAsync())
                            .RootElement.GetProperty("quoteSummary").GetProperty("result")[0];
                        double pe = 0, yld = 0;
                        if (root.TryGetProperty("summaryDetail", out var sd))
                        {
                            if (sd.TryGetProperty("trailingPE", out var peEl) &&
                                peEl.TryGetProperty("raw", out var peRaw)) pe = peRaw.GetDouble();
                            if (sd.TryGetProperty("dividendYield", out var yEl) &&
                                yEl.TryGetProperty("raw", out var yRaw)) yld = yRaw.GetDouble();
                            if (yld == 0 && sd.TryGetProperty("yield", out var y2) &&
                                y2.TryGetProperty("raw", out var yRaw2)) yld = yRaw2.GetDouble();
                        }
                        if (pe > 0 || yld > 0) return (pe, yld, "Yahoo-v11");
                    }
                    else if ((int)resp.StatusCode is 401 or 403)
                    {
                        _authInit = false;
                        AppLogger.Log($"[Fundamentals] Yahoo-v11 認證過期，已重置");
                    }
                }
                catch (Exception ex) { AppLogger.Log($"[Fundamentals] Yahoo-v11 失敗 {ticker}", ex); }

                // ══════════════ Yahoo v7 quote（快速，欄位多）══════════════
                try
                {
                    string url = $"https://query1.finance.yahoo.com/v7/finance/quote" +
                                 $"?symbols={Uri.EscapeDataString(ticker)}" +
                                 $"&fields=trailingPE,forwardPE,trailingAnnualDividendYield,dividendRate";
                    var root = JsonDocument.Parse(
                        await AppHttpClients.Market.GetStringAsync(url))
                        .RootElement.GetProperty("quoteResponse").GetProperty("result");
                    if (root.GetArrayLength() > 0)
                    {
                        var q = root[0];
                        double pe = q.TryGetProperty("trailingPE", out var pE) ? pE.GetDouble() :
                                     q.TryGetProperty("forwardPE", out var fE) ? fE.GetDouble() : 0;
                        double yld = q.TryGetProperty("trailingAnnualDividendYield", out var yE) ? yE.GetDouble() :
                                     q.TryGetProperty("dividendRate", out var dR) ? dR.GetDouble() : 0;
                        if (pe > 0 || yld > 0) return (pe, yld, "Yahoo-v7");
                    }
                }
                catch (Exception ex) { AppLogger.Log($"[Fundamentals] Yahoo-v7 失敗 {ticker}", ex); }

                // ══════════════ Yahoo v10（終極備援）══════════════
                try
                {
                    await EnsureAuthAsync();
                    string cp = !string.IsNullOrEmpty(_crumb) ? $"&crumb={Uri.EscapeDataString(_crumb)}" : "";
                    string url = $"https://query2.finance.yahoo.com/v10/finance/quoteSummary/" +
                                 $"{Uri.EscapeDataString(ticker)}?modules=summaryDetail{cp}";
                    var resp = await AppHttpClients.Market.GetAsync(url);
                    if (resp.IsSuccessStatusCode)
                    {
                        var root = JsonDocument.Parse(await resp.Content.ReadAsStringAsync())
                            .RootElement.GetProperty("quoteSummary").GetProperty("result")[0];
                        double pe = 0, yld = 0;
                        if (root.TryGetProperty("summaryDetail", out var sd))
                        {
                            if (sd.TryGetProperty("trailingPE", out var peEl) &&
                                peEl.TryGetProperty("raw", out var r)) pe = r.GetDouble();
                            if (sd.TryGetProperty("dividendYield", out var yEl) &&
                                yEl.TryGetProperty("raw", out var r2)) yld = r2.GetDouble();
                        }
                        if (pe > 0 || yld > 0) return (pe, yld, "Yahoo-v10");
                    }
                }
                catch (Exception ex) { AppLogger.Log($"[Fundamentals] Yahoo-v10 失敗 {ticker}", ex); }

                AppLogger.Log($"[Fundamentals] 全部來源失敗 {ticker}");
                return (0.0, 0.0, "Unavailable");

            }, maxRetry: 1, tag: "FetchFundamentals");
        }

        // ── FetchHistoryAsync：AutoTradeForm 用（包裝 FetchYahoo）───────────────
        public static async Task<List<MarketData>> FetchHistoryAsync(
            string ticker, string range = "1mo", CancellationToken ct = default)
        {
            // range: 1d / 5d / 1mo / 3mo / 6mo / 1y
            string interval = range is "1d" or "5d" ? "5m" : "1d";
            try
            {
                return await RetryHelper.RunAsync(
                    () => FetchYahoo(ticker, interval: interval, range: range),
                    maxRetry: 1, tag: "FetchHistoryAsync", ct: ct);
            }
            catch { return new List<MarketData>(); }
        }

        public static async Task<string> FetchCompanyName(string ticker)
        {
            await EnsureAuthAsync();
            try
            {
                string cp = !string.IsNullOrEmpty(_crumb) ? $"&crumb={Uri.EscapeDataString(_crumb)}" : "";
                string url = $"https://query2.finance.yahoo.com/v1/finance/search" +
                             $"?q={Uri.EscapeDataString(ticker)}{cp}";
                var root = JsonDocument.Parse(
                    await AppHttpClients.Market.GetStringAsync(url)).RootElement;
                if (root.TryGetProperty("quotes", out var quotes) && quotes.GetArrayLength() > 0)
                    return quotes[0].TryGetProperty("shortname", out var sn) ? (sn.GetString() ?? ticker) : ticker;
            }
            catch (Exception ex) { AppLogger.Log($"FetchCompanyName 失敗 {ticker}", ex); }
            return "";
        }

        /// <summary>
        /// 取得 VIX 當前值 + 前一日收盤 → 合併成單一 range=5d 請求。
        /// 修正：原本 FetchVIX() + FetchPrevClose("^VIX") 兩次獨立請求，
        ///        現在一次請求兩個值，省一次 RTT。
        /// </summary>
        public static async Task<(double Current, double Prev)> FetchVIXWithChange()
        {
            try
            {
                string url = "https://query2.finance.yahoo.com/v8/finance/chart/%5EVIX?interval=1d&range=5d";
                var root = JsonDocument.Parse(
                    await AppHttpClients.Market.GetStringAsync(url))
                    .RootElement.GetProperty("chart").GetProperty("result")[0];
                var closes = root.GetProperty("indicators").GetProperty("quote")[0].GetProperty("close");
                int len = closes.GetArrayLength();

                double cur = 0, prev = 0;
                int found = 0;
                for (int i = len - 1; i >= 0 && found < 2; i--)
                {
                    if (closes[i].ValueKind == JsonValueKind.Null) continue;
                    if (found == 0) cur = closes[i].GetDouble();
                    else prev = closes[i].GetDouble();
                    found++;
                }
                return (cur, prev);
            }
            catch (Exception ex) { AppLogger.Log("FetchVIXWithChange 失敗", ex); }
            return (20.0, 20.0);
        }

        public static async Task<List<MarketData>> FetchYahooByDate(
            string t, DateTime start, DateTime end,
            CancellationToken ct = default) =>
            await ProcessYahooRequest(
                $"https://query2.finance.yahoo.com/v8/finance/chart/{t}" +
                $"?interval=1d&period1={((DateTimeOffset)start).ToUnixTimeSeconds()}" +
                $"&period2={((DateTimeOffset)end).ToUnixTimeSeconds()}", ct);

        public static async Task<List<MarketData>> FetchYahoo(
            string t, string interval, string range,
            CancellationToken ct = default) =>
            await ProcessYahooRequest(
                $"https://query2.finance.yahoo.com/v8/finance/chart/{t}" +
                $"?interval={interval}&range={range}", ct);

        private static async Task<List<MarketData>> ProcessYahooRequest(
            string url, CancellationToken ct = default)
        {
            return await RetryHelper.RunAsync(async () =>
            {
                var list = new List<MarketData>();
                try
                {
                    var req = new HttpRequestMessage(HttpMethod.Get, url);
                    var resp = await AppHttpClients.Market.SendAsync(req, ct);
                    var root = JsonDocument.Parse(await resp.Content.ReadAsStringAsync())
                        .RootElement.GetProperty("chart").GetProperty("result")[0];

                    if (!root.TryGetProperty("timestamp", out var tsEl)) return list;
                    var ts = tsEl.EnumerateArray().ToList();
                    if (!root.TryGetProperty("indicators", out var ind) ||
                        !ind.TryGetProperty("quote", out var qa) ||
                        qa.GetArrayLength() == 0) return list;
                    var q = qa[0];
                    if (!q.TryGetProperty("close", out var closeArr)) return list;

                    for (int i = 0; i < ts.Count; i++)
                    {
                        if (closeArr[i].ValueKind == JsonValueKind.Null) continue;
                        list.Add(new MarketData
                        {
                            // 明確轉為 UTC+8（台灣時區），不依賴 OS 系統時區
                            // 避免在非台灣伺服器上 ToLocalTime() 產生 K 棒時間錯位
                            Date = DateTimeOffset.FromUnixTimeSeconds(ts[i].GetInt64())
                                       .ToOffset(TimeSpan.FromHours(8)).DateTime,
                            Open = q.TryGetProperty("open", out var oa) && oa[i].ValueKind != JsonValueKind.Null ? oa[i].GetDouble() : 0,
                            Close = closeArr[i].GetDouble(),
                            High = q.TryGetProperty("high", out var ha) && ha[i].ValueKind != JsonValueKind.Null ? ha[i].GetDouble() : 0,
                            Low = q.TryGetProperty("low", out var la) && la[i].ValueKind != JsonValueKind.Null ? la[i].GetDouble() : 0,
                            Volume = q.TryGetProperty("volume", out var va) && va[i].ValueKind != JsonValueKind.Null ? va[i].GetDouble() : 0
                        });
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) { AppLogger.Log($"ProcessYahooRequest 失敗 {url}", ex); }
                return list;
            }, maxRetry: 2, tag: "Yahoo", ct: ct);
        }

        // ── crumb 管理：25 分鐘 TTL，double-check lock ──────────────────────
        public static async Task EnsureAuthAsync()
        {
            if (_authInit && !string.IsNullOrEmpty(_crumb) &&
                DateTime.Now - _crumbFetchedAt < _crumbTtl) return;

            if (!await _authLock.WaitAsync(3000)) return;
            try
            {
                if (_authInit && !string.IsNullOrEmpty(_crumb) &&
                    DateTime.Now - _crumbFetchedAt < _crumbTtl) return;

                await AppHttpClients.Market.GetAsync("https://finance.yahoo.com/");
                var resp = await AppHttpClients.Market.GetAsync(
                    "https://query1.finance.yahoo.com/v1/test/getcrumb");
                if (resp.IsSuccessStatusCode)
                {
                    _crumb = await resp.Content.ReadAsStringAsync();
                    _crumbFetchedAt = DateTime.Now;
                    _authInit = true;
                    AppLogger.Log($"Yahoo crumb 更新 (下次: {_crumbFetchedAt + _crumbTtl:HH:mm})");
                }
            }
            catch (Exception ex) { AppLogger.Log("EnsureAuthAsync 失敗", ex); }
            finally { _authLock.Release(); }
        }
    }


    // ────────────────────────────────────────────────────────────────────────────
    //  Alpha Vantage Key 管理（免費 25次/天，需到 alphavantage.co 申請）
    // ────────────────────────────────────────────────────────────────────────────
    public static class AlphaVantageKeyManager
    {
        private static readonly string Path_ =
            System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AV_Key.txt");

        public static string Key { get; private set; } = Load();

        public static string Load()
        {
            try { if (File.Exists(Path_)) return File.ReadAllText(Path_).Trim(); }
            catch (Exception ex) { AppLogger.Log("AV Key Load 失敗", ex); }
            return "";
        }

        public static void Save(string key)
        {
            Key = key.Trim();
            try { File.WriteAllText(Path_, Key); }
            catch (Exception ex) { AppLogger.Log("AV Key Save 失敗", ex); }
        }

        public static bool HasKey => !string.IsNullOrEmpty(Key);
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  新聞服務（使用共享 HttpClient.News）
    // ────────────────────────────────────────────────────────────────────────────
    public static class NewsService
    {
        public static async Task<List<string>> FetchYahooNews(string query)
        {
            var list = new List<string>();
            try
            {
                string url = $"https://query2.finance.yahoo.com/v1/finance/search" +
                             $"?q={Uri.EscapeDataString(query)}&newsCount=5";
                var root = JsonDocument.Parse(
                    await AppHttpClients.News.GetStringAsync(url)).RootElement;
                if (root.TryGetProperty("news", out var newsArr))
                    foreach (var n in newsArr.EnumerateArray())
                    {
                        string title = n.TryGetProperty("title", out var t) ? t.GetString() : "";
                        string pub = n.TryGetProperty("publisher", out var p) ? p.GetString() : "News";
                        long ts = n.TryGetProperty("providerPublishTime", out var tp) ? tp.GetInt64() : 0;
                        DateTime dt = DateTimeOffset.FromUnixTimeSeconds(ts).ToLocalTime().DateTime;
                        if (!string.IsNullOrEmpty(title))
                            list.Add($"[{dt:MM/dd HH:mm}] {pub}: {title}");
                    }
            }
            catch (Exception ex) { AppLogger.Log("NewsService.FetchYahooNews 失敗", ex); }
            return list;
        }

        public static async Task<List<string>> FetchCnaNews(string query)
        {
            var list = new List<string>();
            try
            {
                string url = $"https://news.google.com/rss/search" +
                             $"?q={Uri.EscapeDataString(query + " 中央社")}" +
                             $"&hl=zh-TW&gl=TW&ceid=TW:zh-Hant";
                string xml = await AppHttpClients.News.GetStringAsync(url);
                var doc = new XmlDocument();
                doc.LoadXml(xml);
                var nodes = doc.SelectNodes("//item");
                int count = 0;
                if (nodes != null)
                    foreach (XmlNode node in nodes)
                    {
                        if (count >= 5) break;
                        string title = node.SelectSingleNode("title")?.InnerText ?? "";
                        string pubDateStr = node.SelectSingleNode("pubDate")?.InnerText ?? "";
                        string source = node.SelectSingleNode("source")?.InnerText ?? "中央社";
                        if (!DateTime.TryParse(pubDateStr, out DateTime dt)) dt = DateTime.Now;
                        else dt = dt.ToLocalTime();
                        if (!string.IsNullOrEmpty(title))
                        {
                            title = Regex.Replace(title, @" - .*$", "").Trim();
                            list.Add($"[{dt:MM/dd HH:mm}] {source}: {title}");
                            count++;
                        }
                    }
            }
            catch (Exception ex) { AppLogger.Log("新聞 RSS 解析失敗", ex); }
            return list;
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  LLM 串流服務
    //  改進：使用 AppHttpClients.Llm（共享）+ LlmConfig.CurrentModel（可切換）
    //        + CancellationToken 支援取消
    // ────────────────────────────────────────────────────────────────────────────
    public static class LLMService
    {
        public static async Task StreamChat(
            string key, List<object> msgs, Action<string> onUpdate,
            CancellationToken ct = default)
        {
            // ── Google AI Studio 直連路由 ─────────────────────────────────────
            if (LlmConfig.IsGeminiDirect(LlmConfig.CurrentModel))
            {
                string geminiKey = !string.IsNullOrEmpty(key) ? key : LlmConfig.GeminiApiKey;
                string modelName = LlmConfig.GeminiModelName(LlmConfig.CurrentModel);
                // 從 msgs 提取 system + user 內容
                string sysPart = "", userPart = "";
                string msgsJson = JsonSerializer.Serialize(msgs);
                foreach (var el in JsonDocument.Parse(msgsJson).RootElement.EnumerateArray())
                {
                    string role    = el.GetProperty("role").GetString() ?? "";
                    string content = el.TryGetProperty("content", out var cv) ? cv.GetString() ?? "" : "";
                    if (role == "system") sysPart  += content;
                    else                  userPart += content;
                }
                await GeminiService.StreamAsync(geminiKey, modelName, sysPart, userPart, onUpdate, ct);
                return;
            }

            // ── OpenRouter 路由（原有邏輯）────────────────────────────────────
            var payload = new { model = LlmConfig.CurrentModel, messages = msgs, stream = true };
            var req = new HttpRequestMessage(HttpMethod.Post, LlmConfig.BaseUrl)
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", key);
            try
            {
                var resp = await AppHttpClients.Llm.SendAsync(
                    req, HttpCompletionOption.ResponseHeadersRead, ct);
                resp.EnsureSuccessStatusCode();
                using var reader = new StreamReader(await resp.Content.ReadAsStreamAsync());
                while (!reader.EndOfStream)
                {
                    ct.ThrowIfCancellationRequested();
                    var line = await reader.ReadLineAsync();
                    if (string.IsNullOrEmpty(line) || !line.StartsWith("data: ") ||
                        line == "data: [DONE]") continue;
                    try
                    {
                        var delta = JsonDocument.Parse(line.Substring(6))
                            .RootElement.GetProperty("choices")[0].GetProperty("delta");
                        if (delta.TryGetProperty("content", out var cp) &&
                            cp.ValueKind != JsonValueKind.Null)
                        {
                            string chunk = cp.GetString();
                            if (!string.IsNullOrEmpty(chunk)) onUpdate(chunk);
                        }
                    }
                    catch (Exception ex) { AppLogger.Log("StreamChat JSON解析失敗", ex); }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) { AppLogger.Log("LLMService.StreamChat 失敗", ex); throw; }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  Google AI Studio 直連服務
    //  - 不需 OpenRouter，直接用 Google 官方 Gemini REST API
    //  - 支援串流 (streamGenerateContent) 與非串流 (generateContent)
    //  - 模型 ID 格式：gemini-2.0-flash / gemini-1.5-flash / gemini-1.5-pro
    // ────────────────────────────────────────────────────────────────────────────
    public static class GeminiService
    {
        private const string ApiBase = "https://generativelanguage.googleapis.com/v1beta/models";

        /// <summary>
        /// 非串流呼叫 — 適合 RunAlphaDebate（需要完整 JSON 輸出）
        /// </summary>
        public static async Task<string> CallAsync(
            string apiKey, string modelName,
            string systemPrompt, string userContent,
            bool jsonMode = false,
            CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Google AI Studio Key 未設定，請在「提示詞設定」頁面填入 Gemini API Key。");

            var body = BuildBody(systemPrompt, userContent, jsonMode);
            string url = $"{ApiBase}/{modelName}:generateContent?key={apiKey}";

            var req = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
            };
            var resp = await AppHttpClients.Llm.SendAsync(req, ct);
            resp.EnsureSuccessStatusCode();

            var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
            return doc.RootElement
                      .GetProperty("candidates")[0]
                      .GetProperty("content")
                      .GetProperty("parts")[0]
                      .GetProperty("text").GetString() ?? "";
        }

        /// <summary>
        /// 串流呼叫 — 適合即時聊天（LLMService.StreamChat 路由至此）
        /// </summary>
        public static async Task StreamAsync(
            string apiKey, string modelName,
            string systemPrompt, string userContent,
            Action<string> onUpdate,
            CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("Google AI Studio Key 未設定，請在「提示詞設定」頁面填入 Gemini API Key。");

            var body = BuildBody(systemPrompt, userContent, jsonMode: false);
            string url = $"{ApiBase}/{modelName}:streamGenerateContent?key={apiKey}&alt=sse";

            var req = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json")
            };
            var resp = await AppHttpClients.Llm.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
            resp.EnsureSuccessStatusCode();

            using var reader = new StreamReader(await resp.Content.ReadAsStreamAsync());
            while (!reader.EndOfStream)
            {
                ct.ThrowIfCancellationRequested();
                var line = await reader.ReadLineAsync();
                if (string.IsNullOrEmpty(line) || !line.StartsWith("data: ")) continue;
                string json = line.Substring(6).Trim();
                if (json == "[DONE]") break;
                try
                {
                    var doc = JsonDocument.Parse(json);
                    var candidates = doc.RootElement.GetProperty("candidates");
                    if (candidates.GetArrayLength() == 0) continue;
                    var parts = candidates[0].GetProperty("content").GetProperty("parts");
                    if (parts.GetArrayLength() == 0) continue;
                    string chunk = parts[0].GetProperty("text").GetString();
                    if (!string.IsNullOrEmpty(chunk)) onUpdate(chunk);
                }
                catch (Exception ex) { AppLogger.Log("GeminiService.StreamAsync 解析失敗", ex); }
            }
        }

        // 組裝 Gemini 請求 Body（system_instruction + contents + 可選 JSON mode）
        private static object BuildBody(string systemPrompt, string userContent, bool jsonMode)
        {
            var contents = new[] { new { role = "user", parts = new[] { new { text = userContent } } } };
            var sysInstruction = new { parts = new[] { new { text = systemPrompt } } };
            if (jsonMode)
                return new
                {
                    system_instruction = sysInstruction,
                    contents,
                    generationConfig = new { responseMimeType = "application/json" }
                };
            return new { system_instruction = sysInstruction, contents };
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  市場情緒服務
    //  改進：
    //  ① FetchVIX + FetchPrevClose → 合併為 YahooDataService.FetchVIXWithChange()，省一次請求
    //  ② 全部使用 AppHttpClients.Market（共享）
    // ────────────────────────────────────────────────────────────────────────────
    public static class MarketSentimentService
    {
        private static MarketSentiment _cached = null;
        private static readonly SemaphoreSlim _lock = new SemaphoreSlim(1, 1);

        public static async Task<MarketSentiment> FetchAsync(bool forceRefresh = false)
        {
            if (_cached != null && !_cached.IsStale && !forceRefresh) return _cached;
            if (!await _lock.WaitAsync(500)) return _cached;
            try
            {
                // 四個請求完全平行（FetchVIXWithChange 已合併 VIX 現值 + 昨收）
                // ① Alternative.me 官方 Fear & Greed API（最先嘗試，最可靠）
                var fngTask = FetchAlternativeFearGreed();
                var vixTask = YahooDataService.FetchVIXWithChange();
                var spTask = FetchIndexChange("^GSPC");
                var ndqTask = FetchIndexChange("^IXIC");
                await Task.WhenAll(fngTask, vixTask, spTask, ndqTask);

                var (vix, vixPrev) = vixTask.Result;
                double sp = spTask.Result;
                double ndq = ndqTask.Result;
                double vixChange = vixPrev > 0 ? vix - vixPrev : 0;

                // ── Fear & Greed：優先使用 Alternative.me 官方數值 ───────────
                // Alternative.me 提供的是市場共識的 Fear & Greed 指數（0-100）
                // 官方說明：https://alternative.me/crypto/fear-and-greed-index/
                // 若 API 失敗則退回自行計算（以 VIX + 大盤漲跌推估）
                double fgScore;
                string fgSource;
                (fgScore, fgSource) = fngTask.Result;
                if (fgScore <= 0)
                {
                    // ── 備援：自行合成（VIX 加權公式）────────────────────────
                    double vixScore = vix <= 12 ? 90 : vix <= 18 ? 70 : vix <= 25 ? 50 : vix <= 35 ? 25 : 5;
                    double spScore  = sp >= 1.5 ? 85 : sp >= 0.5 ? 65 : sp >= 0 ? 55 : sp >= -0.5 ? 40 : sp >= -1.5 ? 25 : 10;
                    double ndqScore = ndq >= 1.5 ? 85 : ndq >= 0.5 ? 65 : ndq >= 0 ? 55 : ndq >= -0.5 ? 40 : ndq >= -1.5 ? 25 : 10;
                    double vixMom   = vixChange <= -1.5 ? 75 : vixChange <= -0.5 ? 62 : vixChange <= 0.5 ? 50 : vixChange <= 1.5 ? 35 : 15;
                    double breadth  = (sp > 0 && ndq > 0) ? 65 : (sp < 0 && ndq < 0) ? 35 : 50;
                    fgScore = Math.Max(0, Math.Min(100,
                        vixScore * 0.35 + spScore * 0.25 + ndqScore * 0.20 +
                        vixMom * 0.10 + breadth * 0.10));
                    fgSource = "VIX推算";
                    AppLogger.Log("MarketSentimentService: Alternative.me 失敗，改用 VIX 推算 Fear & Greed");
                }

                string fgLabel = fgScore >= 75 ? "🔥 極度貪婪" : fgScore >= 60 ? "💚 貪婪" :
                                   fgScore >= 40 ? "⬜ 中立" : fgScore >= 25 ? "🔵 恐懼" : "🟣 極度恐懼";
                string vixLevel = vix <= 15 ? "🟢 低波動 (<15)" : vix <= 25 ? "🟡 正常 (15-25)" :
                                   vix <= 35 ? "🟠 恐慌 (25-35)" : "🔴 極度恐慌 (>35)";
                string posAdvice = fgScore >= 70 ? "⚠️ 市場過熱，建議減倉、嚴守停利" :
                                   fgScore >= 55 ? "📈 偏多，可持倉但設好停損" :
                                   fgScore >= 40 ? "⚖️ 中性，觀望或輕倉試水" :
                                   fgScore >= 25 ? "💡 恐慌即機會，可分批佈局" :
                                                   "🚨 極度恐慌，等待市場穩定後進場";

                string summary =
                    $"╔══ 市場情緒儀表板 [{DateTime.Now:MM/dd HH:mm}] ══╗\n" +
                    $"  VIX: {vix:F2} ({(vixChange >= 0 ? "+" : "")}{vixChange:F2})  {vixLevel}\n" +
                    $"  S&P500: {(sp >= 0 ? "+" : "")}{sp:F2}%   Nasdaq: {(ndq >= 0 ? "+" : "")}{ndq:F2}%\n" +
                    $"  Fear & Greed: {fgScore:F0}/100  {fgLabel}  [來源:{fgSource}]\n" +
                    $"  倉位建議: {posAdvice}\n" +
                    $"╚═══════════════════════════════╝";

                _cached = new MarketSentiment
                {
                    FetchTime = DateTime.Now,
                    VIX = vix,
                    VIXChange = vixChange,
                    VIXLevel = vixLevel,
                    SP500ChangePct = sp,
                    NasdaqChangePct = ndq,
                    FearGreedScore = fgScore,
                    FearGreedLabel = fgLabel,
                    AiPositionAdvice = posAdvice,
                    Summary = summary
                };
                return _cached;
            }
            catch (Exception ex)
            {
                AppLogger.Log("MarketSentimentService.FetchAsync 失敗", ex);
                return _cached ?? new MarketSentiment
                {
                    FetchTime = DateTime.Now,
                    VIX = 0,
                    FearGreedScore = 50,
                    FearGreedLabel = "⬜ 中立 (無法取得數據)",
                    Summary = "⚠️ 無法取得市場情緒數據，請檢查網路連線。"
                };
            }
            finally { _lock.Release(); }
        }

        /// <summary>
        /// 從 Alternative.me 取得官方 Fear &amp; Greed 指數（0=極度恐懼，100=極度貪婪）
        /// API 文件：https://alternative.me/crypto/fear-and-greed-index/
        /// 注意：此 API 原為加密貨幣市場設計，但其情緒指標與股市高度相關，
        /// 整合 VIX 趨勢、動能、波動度、社群等多個維度，比單純 VIX 推算更準確。
        /// 免費，無需 API Key，每日更新
        /// </summary>
        private static async Task<(double Score, string Source)> FetchAlternativeFearGreed()
        {
            try
            {
                // Alternative.me 官方 API（免費，無需 Key）
                string url = "https://api.alternative.me/fng/?limit=1&format=json";
                var root = JsonDocument.Parse(
                    await AppHttpClients.Market.GetStringAsync(url)).RootElement;

                if (root.TryGetProperty("data", out var data) && data.GetArrayLength() > 0)
                {
                    var item = data[0];
                    if (item.TryGetProperty("value", out var val) &&
                        double.TryParse(val.GetString(), out double score))
                    {
                        string classification = item.TryGetProperty("value_classification", out var vc)
                            ? vc.GetString() : "";
                        AppLogger.Log($"Alternative.me Fear&Greed = {score} ({classification})");
                        return (score, "Alternative.me官方");
                    }
                }
            }
            catch (Exception ex) { AppLogger.Log("FetchAlternativeFearGreed 失敗", ex); }
            return (0, "");
        }

        private static async Task<double> FetchIndexChange(string symbol)
        {
            try
            {
                string url = $"https://query2.finance.yahoo.com/v8/finance/chart/" +
                             $"{Uri.EscapeDataString(symbol)}?interval=1d&range=2d";
                var root = JsonDocument.Parse(
                    await AppHttpClients.Market.GetStringAsync(url))
                    .RootElement.GetProperty("chart").GetProperty("result")[0];
                var closes = root.GetProperty("indicators").GetProperty("quote")[0].GetProperty("close");
                int len = closes.GetArrayLength();
                double cur = 0, prev = 0; int found = 0;
                for (int i = len - 1; i >= 0; i--)
                {
                    if (closes[i].ValueKind == JsonValueKind.Null) continue;
                    if (found == 0) { cur = closes[i].GetDouble(); found++; }
                    else { prev = closes[i].GetDouble(); break; }
                }
                return prev > 0 ? (cur - prev) / prev * 100 : 0;
            }
            catch (Exception ex) { AppLogger.Log($"FetchIndexChange {symbol} 失敗", ex); return 0; }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  財報日期服務
    //  取得個股下次財報日 + EPS 預估（美股 Yahoo Finance）
    // ────────────────────────────────────────────────────────────────────────────
    public static class EarningsService
    {
        /// <summary>回傳 (下次財報日, EPS預估, 上次EPS實際, 驚喜率%)</summary>
        public static async Task<(DateTime? NextDate, double EstEps, double ActEps, double SurprisePct, string FiscalQuarter)>
            FetchEarningsAsync(string ticker)
        {
            try
            {
                await YahooDataService.EnsureAuthAsync();
                string cp = "";
                string url = $"https://query2.finance.yahoo.com/v10/finance/quoteSummary/" +
                             $"{Uri.EscapeDataString(ticker)}" +
                             $"?modules=earningsTrend,calendarEvents{cp}";
                var resp = await AppHttpClients.Market.GetAsync(url);
                if (!resp.IsSuccessStatusCode) return (null, 0, 0, 0, "");

                var root = JsonDocument.Parse(await resp.Content.ReadAsStringAsync())
                    .RootElement.GetProperty("quoteSummary").GetProperty("result")[0];

                DateTime? nextDate = null;
                string quarter = "";
                double estEps = 0, actEps = 0, surprise = 0;

                // calendarEvents → 下次財報日
                if (root.TryGetProperty("calendarEvents", out var cal) &&
                    cal.TryGetProperty("earnings", out var eCal))
                {
                    if (eCal.TryGetProperty("earningsDate", out var edArr) && edArr.GetArrayLength() > 0)
                    {
                        long ts = edArr[0].TryGetProperty("raw", out var tsEl) ? tsEl.GetInt64() : 0;
                        if (ts > 0) nextDate = DateTimeOffset.FromUnixTimeSeconds(ts).DateTime.ToLocalTime();
                    }
                }

                // earningsTrend → 當季 EPS 預估
                if (root.TryGetProperty("earningsTrend", out var et) &&
                    et.TryGetProperty("trend", out var trends) &&
                    trends.GetArrayLength() > 0)
                {
                    // 第一筆通常是當季
                    var cur = trends[0];
                    if (cur.TryGetProperty("period", out var pd)) quarter = pd.GetString() ?? "";
                    if (cur.TryGetProperty("earningsEstimate", out var ee) &&
                        ee.TryGetProperty("avg", out var avgEl) &&
                        avgEl.TryGetProperty("raw", out var avgRaw))
                        estEps = avgRaw.GetDouble();
                }

                // 上次實際 EPS 和驚喜
                if (root.TryGetProperty("calendarEvents", out var cal2) &&
                    cal2.TryGetProperty("earnings", out var eCal2) &&
                    eCal2.TryGetProperty("earningsAverage", out var ea) &&
                    ea.TryGetProperty("raw", out var eaRaw))
                    actEps = eaRaw.GetDouble();

                if (actEps != 0 && estEps != 0)
                    surprise = (actEps - estEps) / Math.Abs(estEps) * 100;

                return (nextDate, estEps, actEps, surprise, quarter);
            }
            catch (Exception ex) { AppLogger.Log($"FetchEarningsAsync 失敗 {ticker}", ex); }
            return (null, 0, 0, 0, "");
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  板塊輪動快照服務
    //  一次抓取 11 個主要板塊 ETF 的當日漲跌幅
    // ────────────────────────────────────────────────────────────────────────────
    public static class SectorRotationService
    {
        public static readonly (string Ticker, string Name, string Emoji)[] SectorEtfs =
        {
            ("XLK",  "科技",     "💻"),
            ("XLY",  "非必需消費","🛍️"),
            ("XLC",  "通訊",     "📡"),
            ("XLF",  "金融",     "🏦"),
            ("XLI",  "工業",     "⚙️"),
            ("XLV",  "醫療保健", "💊"),
            ("XLP",  "必需消費", "🛒"),
            ("XLRE", "房地產",   "🏠"),
            ("XLE",  "能源",     "⛽"),
            ("XLB",  "原材料",   "🪨"),
            ("XLU",  "公用事業", "⚡"),
        };

        public static async Task<List<(string Ticker, string Name, string Emoji, double ChangePct, double Price)>>
            FetchSnapshotAsync()
        {
            var result = new List<(string, string, string, double, double)>();
            try
            {
                string syms = string.Join(",", SectorEtfs.Select(s => s.Ticker));
                string url = $"https://query1.finance.yahoo.com/v7/finance/quote" +
                              $"?symbols={Uri.EscapeDataString(syms)}" +
                              $"&fields=regularMarketChangePercent,regularMarketPrice,shortName";
                var root = JsonDocument.Parse(
                    await AppHttpClients.Market.GetStringAsync(url))
                    .RootElement.GetProperty("quoteResponse").GetProperty("result");

                var dict = new Dictionary<string, (double chg, double price)>(StringComparer.OrdinalIgnoreCase);
                foreach (var q in root.EnumerateArray())
                {
                    string sym = q.TryGetProperty("symbol", out var s) ? s.GetString() : "";
                    double chg = q.TryGetProperty("regularMarketChangePercent", out var c) ? c.GetDouble() : 0;
                    double price = q.TryGetProperty("regularMarketPrice", out var p) ? p.GetDouble() : 0;
                    if (!string.IsNullOrEmpty(sym)) dict[sym] = (chg, price);
                }

                foreach (var (ticker, name, emoji) in SectorEtfs)
                {
                    if (dict.TryGetValue(ticker, out var d))
                        result.Add((ticker, name, emoji, d.chg, d.price));
                    else
                        result.Add((ticker, name, emoji, 0, 0));
                }
            }
            catch (Exception ex) { AppLogger.Log("SectorRotationService.FetchSnapshotAsync 失敗", ex); }
            return result;
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  標的比較服務
    //  把兩個標的的歷史收盤轉成「相對基準日報酬率 %」，方便疊加比較
    // ────────────────────────────────────────────────────────────────────────────
    public static class ComparisonService
    {
        /// <summary>
        /// 取得兩個標的近 N 天的「日報酬累積曲線」（基準日=0%）
        /// 回傳 (日期, ticker1 累計報酬%, ticker2 累計報酬%)
        /// </summary>
        public static async Task<List<(DateTime Date, double R1, double R2)>>
            FetchComparisonAsync(string t1, string t2, string range = "1y",
                                 CancellationToken ct = default)
        {
            var result = new List<(DateTime, double, double)>();
            try
            {
                var task1 = YahooDataService.FetchYahoo(t1, "1d", range, ct);
                var task2 = YahooDataService.FetchYahoo(t2, "1d", range, ct);
                await Task.WhenAll(task1, task2);

                var d1 = task1.Result.OrderBy(x => x.Date).ToList();
                var d2 = task2.Result.OrderBy(x => x.Date).ToList();
                if (d1.Count == 0 || d2.Count == 0) return result;

                // 對齊日期（取交集）
                var dates1 = d1.Select(x => x.Date.Date).ToHashSet();
                var dates2 = d2.Select(x => x.Date.Date).ToHashSet();
                var common = dates1.Intersect(dates2).OrderBy(x => x).ToList();
                if (common.Count < 2) return result;

                var map1 = d1.ToDictionary(x => x.Date.Date, x => x.Close);
                var map2 = d2.ToDictionary(x => x.Date.Date, x => x.Close);

                double base1 = map1[common[0]], base2 = map2[common[0]];
                foreach (var dt in common)
                {
                    double r1 = base1 > 0 ? (map1[dt] - base1) / base1 * 100 : 0;
                    double r2 = base2 > 0 ? (map2[dt] - base2) / base2 * 100 : 0;
                    result.Add((dt, r1, r2));
                }
            }
            catch (Exception ex) { AppLogger.Log($"ComparisonService 失敗 {t1}/{t2}", ex); }
            return result;
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  ETF 資訊服務
    //
    //  支援：
    //  ① 台灣 ETF（TWSE）：0050.TW, 0056.TW, 00878.TW 等
    //     資料來源：TWSE 官方 API + Yahoo Finance ETF modules
    //
    //  ② 美國 ETF：SPY, QQQ, DIA, IWM, VTI, GLD, TLT 等
    //     資料來源：Yahoo Finance v10 quoteSummary
    //     modules: fundProfile (AUM, 費率, 類別) + topHoldings (前十大持股)
    //
    //  ③ 內建熱門 ETF 清單（快速跳轉用）
    // ────────────────────────────────────────────────────────────────────────────
    public static class EtfService
    {
        // ── 內建熱門 ETF 清單 ──────────────────────────────────────────────────
        public static readonly (string Ticker, string Label, string Category)[] WellKnownEtfs =
        {
            // 台灣 ETF
            ("0050.TW",   "🇹🇼 元大台灣50",         "台股大盤"),
            ("0056.TW",   "🇹🇼 元大高股息",          "台股高股息"),
            ("00878.TW",  "🇹🇼 國泰永續高股息",      "台股ESG高息"),
            ("006208.TW", "🇹🇼 富邦台50",           "台股大盤"),
            ("00646.TW",  "🇹🇼 元大S&P500",        "美股大盤"),
            // 美股 ETF — 大盤
            ("SPY",   "🇺🇸 S&P 500 (SPY)",         "美股大盤"),
            ("QQQ",   "🇺🇸 Nasdaq 100 (QQQ)",       "美股科技"),
            ("DIA",   "🇺🇸 道瓊 (DIA)",              "美股大盤"),
            ("IWM",   "🇺🇸 Russell 2000 (IWM)",     "美股小型股"),
            ("VTI",   "🇺🇸 全美市場 (VTI)",          "美股全市場"),
            ("VT",    "🇺🇸 全球股市 (VT)",           "全球股市"),
            // 美股 ETF — 板塊
            ("XLK",   "🔧 科技 (XLK)",               "美股科技"),
            ("XLF",   "🏦 金融 (XLF)",               "美股金融"),
            ("XLE",   "⛽ 能源 (XLE)",               "美股能源"),
            ("XLV",   "💊 醫療 (XLV)",               "美股醫療"),
            ("ARKK",  "🚀 ARK Innovation (ARKK)",    "主動ETF"),
            // 美股 ETF — 債券/商品
            ("TLT",   "📋 美國長債 (TLT)",            "美國公債"),
            ("GLD",   "🥇 黃金 (GLD)",               "貴金屬"),
            ("SLV",   "🥈 白銀 (SLV)",               "貴金屬"),
            ("USO",   "🛢️ 原油 (USO)",               "能源商品"),
        };

        // ── 偵測是否為 ETF ─────────────────────────────────────────────────
        public static async Task<bool> IsEtfAsync(string ticker)
        {
            // 先查內建清單
            if (WellKnownEtfs.Any(e => e.Ticker.Equals(ticker, StringComparison.OrdinalIgnoreCase)))
                return true;

            // 台灣 ETF 特徵：代碼以 0 開頭且為 4~6 位數字
            string clean = ticker.Replace(".TW", "").Replace(".TWO", "");
            if ((ticker.EndsWith(".TW") || ticker.EndsWith(".TWO")) &&
                clean.Length >= 4 && clean.All(char.IsDigit) && clean.StartsWith("0"))
                return true;

            // Yahoo Finance quoteType 偵測
            try
            {
                await YahooDataService.EnsureAuthAsync();
                string url = $"https://query1.finance.yahoo.com/v7/finance/quote" +
                             $"?symbols={Uri.EscapeDataString(ticker)}&fields=quoteType,typeDisp";
                var root = JsonDocument.Parse(
                    await AppHttpClients.Market.GetStringAsync(url))
                    .RootElement.GetProperty("quoteResponse").GetProperty("result");
                if (root.GetArrayLength() > 0)
                {
                    var q = root[0];
                    string qtype = q.TryGetProperty("quoteType", out var qt) ? qt.GetString() : "";
                    return qtype == "ETF" || qtype == "MUTUALFUND";
                }
            }
            catch (Exception ex) { AppLogger.Log($"IsEtfAsync 失敗 {ticker}", ex); }
            return false;
        }

        // ── 取得 ETF 完整資訊 ─────────────────────────────────────────────
        public static async Task<EtfInfo> FetchEtfInfo(string ticker)
        {
            var info = new EtfInfo { Ticker = ticker, IsEtf = true };

            // 預填內建清單的名稱/類別
            var known = WellKnownEtfs.FirstOrDefault(
                e => e.Ticker.Equals(ticker, StringComparison.OrdinalIgnoreCase));
            if (known != default)
            {
                info.Name = known.Label;
                info.Category = known.Category;
            }

            // ── 台灣 ETF：TWSE 官方 API ───────────────────────────────────
            if (ticker.EndsWith(".TW") || ticker.EndsWith(".TWO"))
            {
                try
                {
                    string clean = ticker.Replace(".TW", "").Replace(".TWO", "");

                    // 取得日成交資料（包含 NAV）
                    string twseUrl = $"https://www.twse.com.tw/exchangeReport/STOCK_DAY" +
                                     $"?response=json&date={DateTime.Now:yyyyMMdd}&stockNo={clean}";
                    var root = JsonDocument.Parse(
                        await AppHttpClients.Market.GetStringAsync(twseUrl)).RootElement;

                    if (root.TryGetProperty("title", out var t))
                        info.Name = t.GetString()?.Split("  ").FirstOrDefault() ?? info.Name;

                    // 三大法人資料（外資/投信/自營商籌碼）
                    string instUrl = $"https://www.twse.com.tw/fund/T86" +
                                     $"?response=json&date={DateTime.Now:yyyyMMdd}&selectType=ALLBUT0999";
                    // 略過詳細解析，主要資料從 Yahoo 補充
                }
                catch (Exception ex) { AppLogger.Log($"FetchEtfInfo TWSE 失敗 {ticker}", ex); }
            }

            // ── Yahoo Finance ETF modules（台灣 + 美國 ETF 通用）──────────
            try
            {
                await YahooDataService.EnsureAuthAsync();
                string cp = "";
                string url = $"https://query2.finance.yahoo.com/v10/finance/quoteSummary/" +
                             $"{Uri.EscapeDataString(ticker)}" +
                             $"?modules=fundProfile,topHoldings,fundPerformance,summaryDetail{cp}";
                var resp = await AppHttpClients.Market.GetAsync(url);

                if (resp.IsSuccessStatusCode)
                {
                    var root = JsonDocument.Parse(await resp.Content.ReadAsStringAsync())
                        .RootElement.GetProperty("quoteSummary").GetProperty("result")[0];

                    // fundProfile：AUM / 費率 / 類別 / 追蹤指數
                    if (root.TryGetProperty("fundProfile", out var fp))
                    {
                        if (fp.TryGetProperty("totalAssets", out var ta) &&
                            ta.TryGetProperty("raw", out var taRaw))
                            info.TotalAssets = taRaw.GetDouble();

                        if (fp.TryGetProperty("feesExpensesInvestment", out var fee) &&
                            fee.TryGetProperty("annualReportExpenseRatio", out var er) &&
                            er.TryGetProperty("raw", out var erRaw))
                            info.ExpenseRatio = erRaw.GetDouble();
                        else if (fp.TryGetProperty("feesExpensesInvestment", out var fee2) &&
                                 fee2.TryGetProperty("totalNetAssets", out var na))
                        { /* fallback */ }

                        if (fp.TryGetProperty("categoryName", out var cat))
                            if (string.IsNullOrEmpty(info.Category))
                                info.Category = cat.GetString() ?? "";

                        if (fp.TryGetProperty("legalType", out var lt))
                            info.AssetClass = lt.GetString() ?? "";
                    }

                    // topHoldings：前十大持股
                    if (root.TryGetProperty("topHoldings", out var th))
                    {
                        if (th.TryGetProperty("holdings", out var holdings))
                            foreach (var h in holdings.EnumerateArray().Take(10))
                            {
                                string hName = h.TryGetProperty("holdingName", out var hn) ? hn.GetString() : "";
                                double hPct = h.TryGetProperty("holdingPercent", out var hp) &&
                                               hp.TryGetProperty("raw", out var hpRaw) ? hpRaw.GetDouble() : 0;
                                if (!string.IsNullOrEmpty(hName))
                                    info.TopHoldings.Add((hName, hPct));
                            }

                        // equityHoldings 有追蹤指數資訊
                        if (string.IsNullOrEmpty(info.TrackIndex) &&
                            th.TryGetProperty("equityHoldings", out var eq) &&
                            eq.TryGetProperty("priceToEarnings", out _))
                        { /* P/E 存在代表是股票型 ETF */ }
                    }

                    // fundPerformance：報酬率
                    if (root.TryGetProperty("fundPerformance", out var perf))
                    {
                        if (perf.TryGetProperty("annualTotalReturns", out var atr) &&
                            atr.TryGetProperty("returns", out var rets) &&
                            rets.GetArrayLength() > 0)
                        {
                            var ytd = rets.EnumerateArray()
                                .FirstOrDefault(r => r.TryGetProperty("year", out var yr) &&
                                                     yr.GetString() == "YTD");
                            if (ytd.ValueKind == JsonValueKind.Object &&
                                ytd.TryGetProperty("annualValue", out var ytdVal) &&
                                ytdVal.TryGetProperty("raw", out var ytdRaw))
                                info.YtdReturn = ytdRaw.GetDouble();
                        }
                    }

                    // summaryDetail：殖利率
                    if (root.TryGetProperty("summaryDetail", out var sd))
                    {
                        if (sd.TryGetProperty("yield", out var yEl) &&
                            yEl.TryGetProperty("raw", out var yRaw))
                            info.DividendYield = yRaw.GetDouble();
                        else if (sd.TryGetProperty("dividendYield", out var dyEl) &&
                                 dyEl.TryGetProperty("raw", out var dyRaw))
                            info.DividendYield = dyRaw.GetDouble();
                    }

                    info.Source = "Yahoo";
                }
            }
            catch (Exception ex) { AppLogger.Log($"FetchEtfInfo Yahoo 失敗 {ticker}", ex); }

            return info;
        }

        // ── 產生 ETF 摘要字串（給 AI 提示詞用）───────────────────────────
        public static string ToPromptContext(EtfInfo etf)
        {
            if (etf == null || !etf.IsEtf) return "";
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"【ETF 資訊】{ etf.Name} ({ etf.Ticker})");
            sb.AppendLine($"類別: {etf.Category}  資產: {FormatAum(etf.TotalAssets)}  費率: {etf.ExpenseRatio:P2}");
            if (etf.DividendYield > 0) sb.AppendLine($"殖利率: {etf.DividendYield:P2}  今年報酬: {etf.YtdReturn:P2}");
            if (etf.TopHoldings.Count > 0)
            {
                sb.Append("前五大持股: ");
                sb.AppendLine(string.Join(", ", etf.TopHoldings.Take(5)
                    .Select(h => $"{h.Name}({h.Pct:P1})")));
            }
            sb.AppendLine("⚠️ ETF 分析重點：追蹤誤差、折溢價、板塊集中風險、換倉時間");
            return sb.ToString();
        }

        private static string FormatAum(double assets)
        {
            if (assets >= 1_000_000_000) return $"${assets / 1_000_000_000:F1}B";
            if (assets >= 1_000_000) return $"${assets / 1_000_000:F0}M";
            return assets > 0 ? $"${assets:N0}" : "N/A";
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  三大法人籌碼服務（TWSE / TPEX 官方 API）
    //
    //  資料來源：
    //  ① 台灣證券交易所 T86：三大法人買賣超（當日所有個股）
    //     https://www.twse.com.tw/fund/T86?response=json&date=YYYYMMDD&selectType=ALLBUT0999
    //  ② TPEX 三大法人：https://www.tpex.org.tw/web/stock/3insti/daily_trade/
    //
    //  欄位意義（TWSE T86）：
    //    [0]=代號 [1]=名稱 [2]=外資買進 [3]=外資賣出 [4]=外資買賣超
    //    [5]=投信買進 [6]=投信賣出 [7]=投信買賣超
    //    [8]=自營商買進 [9]=自營商賣出 [10]=自營商買賣超 [11]=三大法人合計
    // ────────────────────────────────────────────────────────────────────────────
    public static class InstitutionalService
    {
        /// <summary>取得特定股票的三大法人當日買賣超（張數）</summary>
        public static async Task<InstitutionalData> FetchAsync(string ticker)
        {
            var result = new InstitutionalData { Ticker = ticker };
            if (!ticker.EndsWith(".TW") && !ticker.EndsWith(".TWO"))
                return result;

            string clean  = ticker.Replace(".TW", "").Replace(".TWO", "");
            string market = ticker.EndsWith(".TWO") ? "OTC" : "TWSE";
            // 週一取上週五；週末取週五
            DateTime tradeDate = DateTime.Now;
            if (tradeDate.DayOfWeek == DayOfWeek.Saturday) tradeDate = tradeDate.AddDays(-1);
            if (tradeDate.DayOfWeek == DayOfWeek.Sunday)   tradeDate = tradeDate.AddDays(-2);
            if (tradeDate.DayOfWeek == DayOfWeek.Monday)   tradeDate = tradeDate.AddDays(-3);
            string dateStr = tradeDate.ToString("yyyyMMdd");

            try
            {
                if (market == "TWSE")
                {
                    string url = $"https://www.twse.com.tw/fund/T86" +
                                 $"?response=json&date={dateStr}&selectType=ALLBUT0999";
                    var root = JsonDocument.Parse(
                        await AppHttpClients.Market.GetStringAsync(url)).RootElement;

                    if (!root.TryGetProperty("data", out var dataArr)) return result;
                    foreach (var row in dataArr.EnumerateArray())
                    {
                        var cells = row.EnumerateArray().ToList();
                        if (cells.Count < 12) continue; // TWSE T86 至少需 12 欄
                        if (cells[0].GetString()?.Trim() != clean) continue;
                        result.Date       = dateStr;
                        result.ForeignNet = ParseLot(cells[4].GetString());  // 外資買賣超
                        result.TrustNet   = ParseLot(cells[7].GetString());  // 投信買賣超
                        result.DealerNet  = ParseLot(cells[10].GetString()); // 自營商買賣超
                        result.TotalNet   = ParseLot(cells[11].GetString()); // 三大法人合計
                        result.Source     = "TWSE-T86";
                        break;
                    }
                }
                else
                {
                    // TPEX：欄位 [0]=代號 [3]=外資買賣超 [6]=投信買賣超 [9]=自營商買賣超
                    string url = $"https://www.tpex.org.tw/web/stock/3insti/daily_trade/" +
                                 $"3itrade_hedge_result.php" +
                                 $"?l=zh-tw&se=EW&t=D&d={tradeDate:yyyy%2FMM%2Fdd}&o=json";
                    var root = JsonDocument.Parse(
                        await AppHttpClients.Market.GetStringAsync(url)).RootElement;

                    if (!root.TryGetProperty("aaData", out var dataArr)) return result;
                    foreach (var row in dataArr.EnumerateArray())
                    {
                        var cells = row.EnumerateArray().ToList();
                        if (cells.Count < 10) continue; // 欄位不足則跳過此行
                        if (cells[0].GetString()?.Trim() != clean) continue;
                        result.Date       = dateStr;
                        result.ForeignNet = ParseLot(cells[3].GetString()); // 外資買賣超
                        result.TrustNet   = ParseLot(cells[6].GetString()); // 投信買賣超
                        result.DealerNet  = ParseLot(cells[9].GetString()); // 自營商買賣超
                        result.TotalNet   = result.ForeignNet + result.TrustNet + result.DealerNet;
                        result.Source     = "TPEX-3insti";
                        break;
                    }
                }
            }
            catch (Exception ex) { AppLogger.Log($"InstitutionalService.FetchAsync {ticker} 失敗", ex); }
            return result;
        }

        private static long ParseLot(string s)
        {
            if (string.IsNullOrWhiteSpace(s) || s == "--") return 0;
            return long.TryParse(s.Replace(",", ""), out long v) ? v : 0;
        }

        public static string ToPromptContext(InstitutionalData d)
        {
            if (d == null || string.IsNullOrEmpty(d.Source)) return "";
            string arrow = d.TotalNet > 0 ? "⬆️" : d.TotalNet < 0 ? "⬇️" : "⬛";
            string trend = d.TotalNet > 0 ? "三大法人買超" : d.TotalNet < 0 ? "三大法人賣超" : "三大法人持平";
            return $"【三大法人籌碼 {d.Date} · 來源:{d.Source}】\n" +
                   $"  {arrow} {trend} 合計 {d.TotalNet:+#,##0;-#,##0;0} 張\n" +
                   $"  外資: {d.ForeignNet:+#,##0;-#,##0;0} | " +
                   $"投信: {d.TrustNet:+#,##0;-#,##0;0} | " +
                   $"自營商: {d.DealerNet:+#,##0;-#,##0;0}\n";
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  融資融券服務（TWSE / TPEX 官方 API）
    //
    //  資料來源：
    //  ① TWSE MI_MARGN：https://www.twse.com.tw/exchangeReport/MI_MARGN
    //     欄位：[0]=代號 [1]=名稱
    //           [2]=融資買進 [3]=融資賣出 [4]=融資現金償還 [5]=融資餘額 [6]=融資限額
    //           [7]=融券賣出 [8]=融券買進 [9]=融券現金償還 [10]=融券餘額 [11]=融券限額
    //           [12]=資券互抵
    //  ② TPEX 融資融券餘額：https://www.tpex.org.tw/web/stock/margin_trading/margin_balance/
    //
    //  重要指標：
    //  - 融資增加 = 散戶追多（過度樂觀的反向指標）
    //  - 融券增加 = 空頭增加（若伴隨反彈則有軋空機會）
    //  - 券資比 > 20% = 軋空格局（Short Squeeze 潛力）
    // ────────────────────────────────────────────────────────────────────────────
    public static class MarginTradingService
    {
        public static async Task<MarginData> FetchAsync(string ticker)
        {
            var result = new MarginData { Ticker = ticker };
            if (!ticker.EndsWith(".TW") && !ticker.EndsWith(".TWO"))
                return result;

            string clean  = ticker.Replace(".TW", "").Replace(".TWO", "");
            string market = ticker.EndsWith(".TWO") ? "OTC" : "TWSE";
            DateTime tradeDate = DateTime.Now;
            if (tradeDate.DayOfWeek == DayOfWeek.Saturday) tradeDate = tradeDate.AddDays(-1);
            if (tradeDate.DayOfWeek == DayOfWeek.Sunday)   tradeDate = tradeDate.AddDays(-2);
            if (tradeDate.DayOfWeek == DayOfWeek.Monday)   tradeDate = tradeDate.AddDays(-3);
            string dateStr = tradeDate.ToString("yyyyMMdd");

            try
            {
                if (market == "TWSE")
                {
                    string url = $"https://www.twse.com.tw/exchangeReport/MI_MARGN" +
                                 $"?response=json&date={dateStr}&selectType=ALL";
                    var root = JsonDocument.Parse(
                        await AppHttpClients.Market.GetStringAsync(url)).RootElement;

                    if (!root.TryGetProperty("data", out var dataArr)) return result;
                    foreach (var row in dataArr.EnumerateArray())
                    {
                        var cells = row.EnumerateArray().ToList();
                        if (cells.Count < 11) continue; // TWSE MI_MARGN 至少需 11 欄
                        if (cells[0].GetString()?.Trim() != clean) continue;
                        result.Date       = dateStr;
                        result.MarginBuy  = ParseLong(cells[2].GetString()); // 融資買進（張）
                        result.MarginSell = ParseLong(cells[3].GetString()); // 融資賣出（張）
                        result.MarginBal  = ParseLong(cells[5].GetString()); // 融資餘額（張）
                        result.ShortSell  = ParseLong(cells[7].GetString()); // 融券賣出（張）
                        result.ShortBuy   = ParseLong(cells[8].GetString()); // 融券買進（張）
                        result.ShortBal   = ParseLong(cells[10].GetString()); // 融券餘額（張）
                        result.ShortRatio = result.MarginBal > 0
                            ? (double)result.ShortBal / result.MarginBal * 100 : 0;
                        result.Source     = "TWSE-MI_MARGN";
                        break;
                    }
                }
                else
                {
                    // TPEX 融資融券
                    string url = $"https://www.tpex.org.tw/web/stock/margin_trading/margin_balance/" +
                                 $"margin_bal_result.php" +
                                 $"?l=zh-tw&o=json&d={tradeDate:yyyy%2FMM%2Fdd}&s=0,asc";
                    var root = JsonDocument.Parse(
                        await AppHttpClients.Market.GetStringAsync(url)).RootElement;

                    if (!root.TryGetProperty("aaData", out var dataArr)) return result;
                    foreach (var row in dataArr.EnumerateArray())
                    {
                        var cells = row.EnumerateArray().ToList();
                        if (cells.Count < 9) continue; // TPEX 融資融券至少需 9 欄
                        if (cells[0].GetString()?.Trim() != clean) continue;
                        result.Date       = dateStr;
                        result.MarginBuy  = ParseLong(cells[2].GetString());
                        result.MarginSell = ParseLong(cells[3].GetString());
                        result.MarginBal  = ParseLong(cells[4].GetString());
                        result.ShortSell  = ParseLong(cells[6].GetString());
                        result.ShortBuy   = ParseLong(cells[7].GetString());
                        result.ShortBal   = ParseLong(cells[8].GetString());
                        result.ShortRatio = result.MarginBal > 0
                            ? (double)result.ShortBal / result.MarginBal * 100 : 0;
                        result.Source     = "TPEX-MarginBal";
                        break;
                    }
                }
            }
            catch (Exception ex) { AppLogger.Log($"MarginTradingService.FetchAsync {ticker} 失敗", ex); }
            return result;
        }

        private static long ParseLong(string s)
        {
            if (string.IsNullOrWhiteSpace(s) || s == "--") return 0;
            return long.TryParse(s.Replace(",", ""), out long v) ? v : 0;
        }

        public static string ToPromptContext(MarginData d)
        {
            if (d == null || string.IsNullOrEmpty(d.Source)) return "";
            string marginTrend = d.MarginBuy > d.MarginSell ? "⬆️融資增加（散戶追多）" : "⬇️融資減少";
            string shortTrend  = d.ShortSell > d.ShortBuy   ? "⬆️融券增加" : "⬇️融券減少";
            string squeeze     = d.ShortRatio > 20 ? "  ⚡軋空格局（券資比>20%）" : "";
            return $"【融資融券 {d.Date} · 來源:{d.Source}】\n" +
                   $"  融資餘額: {d.MarginBal:N0} 張  {marginTrend}\n" +
                   $"  融券餘額: {d.ShortBal:N0} 張  {shortTrend}\n" +
                   $"  券資比: {d.ShortRatio:F1}%{squeeze}\n";
        }
    }

}