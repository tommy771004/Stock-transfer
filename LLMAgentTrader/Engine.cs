using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace LLMAgentTrader
{
    // ────────────────────────────────────────────────────────────────────────────
    //  全域錯誤日誌
    // ────────────────────────────────────────────────────────────────────────────
    public static class AppLogger
    {
        private static readonly string LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AppError.log");
        private static readonly object _lock = new object();

        public static void Log(string message, Exception ex = null)
        {
            try
            {
                string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
                if (ex != null) line += $" | {ex.GetType().Name}: {ex.Message}";
                lock (_lock) { File.AppendAllText(LogPath, line + Environment.NewLine); }
                System.Diagnostics.Debug.WriteLine(line);
            }
            catch { }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  API Key 加密管理 (DPAPI)
    // ────────────────────────────────────────────────────────────────────────────
    public static class ApiKeyManager
    {
        private static readonly string EncPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "APIKey.enc");
        private static readonly string LegacyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "APIKey.txt");

        public static string Load()
        {
            if (File.Exists(EncPath))
            {
                try
                {
                    byte[] enc = File.ReadAllBytes(EncPath);
                    byte[] dec = ProtectedData.Unprotect(enc, null, DataProtectionScope.CurrentUser);
                    return Encoding.UTF8.GetString(dec);
                }
                catch (Exception ex) { AppLogger.Log("ApiKeyManager.Load 解密失敗", ex); }
            }
            if (File.Exists(LegacyPath))
            {
                try
                {
                    string key = File.ReadAllText(LegacyPath).Trim();
                    if (!string.IsNullOrEmpty(key)) { Save(key); File.Delete(LegacyPath); return key; }
                }
                catch (Exception ex) { AppLogger.Log("ApiKeyManager.Migrate 失敗", ex); }
            }
            return "";
        }

        public static void Save(string key)
        {
            try
            {
                byte[] data = Encoding.UTF8.GetBytes(key);
                byte[] enc = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(EncPath, enc);
            }
            catch (Exception ex) { AppLogger.Log("ApiKeyManager.Save 失敗", ex); }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  支撐壓力引擎
    // ────────────────────────────────────────────────────────────────────────────
    public static class SupportResistanceEngine
    {
        public static void CalculatePivots(List<MarketData> list)
        {
            if (list.Count < 20) return;
            int lookback = 20;
            for (int i = lookback; i < list.Count; i++)
            {
                var window = list.Skip(i - lookback).Take(lookback).ToList();
                list[i].ResistanceLevel = window.Max(x => x.High);
                list[i].SupportLevel = window.Min(x => x.Low);
            }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  技術指標引擎
    // ────────────────────────────────────────────────────────────────────────────
    public static class IndicatorEngine
    {
        public static void CalculateAll(List<MarketData> list)
        {
            if (list.Count < 30) return;

            // VWAP
            double cumVol = 0, cumPv = 0;
            for (int i = 0; i < list.Count; i++)
            {
                cumVol += list[i].Volume;
                cumPv += ((list[i].High + list[i].Low + list[i].Close) / 3.0) * list[i].Volume;
                list[i].VWAP = cumVol > 0 ? cumPv / cumVol : list[i].Close;
            }

            // EMA 50
            if (list.Count >= 50)
            {
                double seed = list.Take(50).Average(x => x.Close), ema = seed;
                list[49].EMA_50 = seed;
                for (int i = 50; i < list.Count; i++)
                { ema = (list[i].Close - ema) * (2.0 / 51) + ema; list[i].EMA_50 = ema; }
            }

            // EMA 200
            if (list.Count >= 200)
            {
                double seed = list.Take(200).Average(x => x.Close), ema = seed;
                list[199].EMA_200 = seed;
                for (int i = 200; i < list.Count; i++)
                { ema = (list[i].Close - ema) * (2.0 / 201) + ema; list[i].EMA_200 = ema; }
            }

            // MACD (DIF / DEA / 柱)
            {
                double e12 = list[0].Close, e26 = list[0].Close, sig = 0;
                int sigWarm = 0; double sigSum = 0; bool sigSeeded = false;
                for (int i = 0; i < list.Count; i++)
                {
                    e12 = (list[i].Close - e12) * (2.0 / 13) + e12;
                    e26 = (list[i].Close - e26) * (2.0 / 27) + e26;
                    double macd = e12 - e26;
                    list[i].MACD = macd;
                    if (!sigSeeded)
                    {
                        sigSum += macd; sigWarm++;
                        if (sigWarm == 9) { sig = sigSum / 9.0; sigSeeded = true; list[i].MACD_Signal = sig; list[i].MACD_Hist = macd - sig; }
                    }
                    else
                    { sig = (macd - sig) * (2.0 / 10) + sig; list[i].MACD_Signal = sig; list[i].MACD_Hist = macd - sig; }
                }
            }

            // RSI 14 — Wilder SMMA
            if (list.Count > 14)
            {
                double sumG = 0, sumL = 0;
                for (int i = 1; i <= 14; i++)
                {
                    double d = list[i].Close - list[i - 1].Close;
                    sumG += Math.Max(0, d); sumL += Math.Max(0, -d);
                }
                double avgG = sumG / 14, avgL = sumL / 14;
                list[14].RSI = avgL == 0 ? 100 : 100 - 100.0 / (1 + avgG / avgL);
                for (int i = 15; i < list.Count; i++)
                {
                    double d = list[i].Close - list[i - 1].Close;
                    avgG = (avgG * 13 + Math.Max(0, d)) / 14;
                    avgL = (avgL * 13 + Math.Max(0, -d)) / 14;
                    list[i].RSI = avgL == 0 ? 100 : 100 - 100.0 / (1 + avgG / avgL);
                }
            }

            // ATR + 布林通道 (樣本 std) + K線型態
            double atr = 0;
            for (int i = 1; i < list.Count; i++)
            {
                double tr = Math.Max(list[i].High - list[i].Low,
                            Math.Max(Math.Abs(list[i].High - list[i - 1].Close),
                                     Math.Abs(list[i].Low - list[i - 1].Close)));
                atr = (atr * 13 + tr) / 14;
                list[i].ATR = atr;

                if (i >= 20)
                {
                    var sl = list.Skip(i - 19).Take(20).Select(x => x.Close).ToList();
                    double avg = sl.Average(), sd = Math.Sqrt(sl.Select(x => Math.Pow(x - avg, 2)).Sum() / 19.0);
                    list[i].BB_Middle = avg; list[i].BB_Upper = avg + sd * 2; list[i].BB_Lower = avg - sd * 2;
                    list[i].BB_Width = avg > 0 ? (list[i].BB_Upper - list[i].BB_Lower) / avg * 100 : 0;
                }

                double body = Math.Abs(list[i].Open - list[i].Close), range = list[i].High - list[i].Low;
                if (range > 0 && body / range < 0.1) list[i].Pattern = "十字星";
                else if (list[i].Close > list[i].Open && list[i - 1].Close < list[i - 1].Open && list[i].Close > list[i - 1].Open) list[i].Pattern = "多頭吞噬";
                else if (list[i].Close < list[i].Open && list[i - 1].Close > list[i - 1].Open && list[i].Close < list[i - 1].Open) list[i].Pattern = "空頭吞噬";
                else list[i].Pattern = "-";
            }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  Fibonacci 回調引擎
    // ────────────────────────────────────────────────────────────────────────────
    public static class FibonacciEngine
    {
        private static readonly double[] Ratios = { 0.0, 0.236, 0.382, 0.5, 0.618, 0.786, 1.0 };

        /// <summary>
        /// 計算最近 lookback 根 K 棒的 Fibonacci 回調層級（自動識別上升/下降趨勢）
        /// </summary>
        public static List<FibLevel> Calculate(List<MarketData> data, int lookback = 100)
        {
            var levels = new List<FibLevel>();
            if (data == null || data.Count < 10) return levels;

            var window = data.TakeLast(Math.Min(lookback, data.Count)).ToList();
            double swingHigh = window.Max(x => x.High);
            double swingLow = window.Min(x => x.Low);
            if (swingHigh <= swingLow) return levels;

            double last = window.Last().Close;
            // 判斷目前是回調 (從高點跌下) 還是反彈 (從低點漲上)
            bool isRetracement = last < swingHigh && (swingHigh - last) > (last - swingLow);

            foreach (var ratio in Ratios)
            {
                double price;
                if (isRetracement)
                    // 從高點計算回調: 100% = 高點, 0% = 低點
                    price = swingHigh - (swingHigh - swingLow) * ratio;
                else
                    // 從低點計算反彈: 0% = 低點, 100% = 高點
                    price = swingLow + (swingHigh - swingLow) * ratio;

                levels.Add(new FibLevel
                {
                    Ratio = ratio,
                    Price = price,
                    Label = $"Fib {ratio:P1}"
                });
            }
            return levels;
        }

        /// <summary>取得最近一次波段的高低點</summary>
        public static (double High, double Low, int HighIdx, int LowIdx) GetSwingPoints(List<MarketData> data, int lookback = 100)
        {
            if (data == null || data.Count < 2) return (0, 0, 0, 0);
            var window = data.TakeLast(Math.Min(lookback, data.Count)).ToList();
            int offset = data.Count - window.Count;
            double sh = window.Max(x => x.High), sl = window.Min(x => x.Low);
            int hiIdx = offset + window.FindIndex(x => x.High == sh);
            int loIdx = offset + window.FindIndex(x => x.Low == sl);
            return (sh, sl, hiIdx, loIdx);
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  多時間框架分析引擎
    // ────────────────────────────────────────────────────────────────────────────
    public static class MultiTimeframeEngine
    {
        /// <summary>分析三個時間框架並回傳共振信號</summary>
        public static async Task<MultiTimeframeSignal> AnalyzeAsync(
            string ticker, CancellationToken ct = default)
        {
            var signal = new MultiTimeframeSignal { Ticker = ticker };
            try
            {
                // 並行拉取三個時間框架
                var taskWeekly = YahooDataService.FetchYahoo(ticker, "1wk", "2y", ct);
                var taskDaily = YahooDataService.FetchYahoo(ticker, "1d", "6mo", ct);
                var taskHourly = YahooDataService.FetchYahoo(ticker, "1h", "5d", ct);

                await Task.WhenAll(taskWeekly, taskDaily, taskHourly);

                var weekly = taskWeekly.Result;
                var daily = taskDaily.Result;
                var hourly = taskHourly.Result;

                // 週線
                if (weekly.Count >= 30)
                {
                    IndicatorEngine.CalculateAll(weekly);
                    var wLast = weekly.Last();
                    signal.Weekly_RSI = wLast.RSI;
                    signal.Weekly_MACD_Hist = wLast.MACD_Hist;
                    signal.Weekly_Pattern = wLast.Pattern;
                    signal.Weekly_Trend = wLast.EMA_50 > 0
                        ? (wLast.Close > wLast.EMA_50 ? "多頭" : "空頭")
                        : "-";
                }

                // 日線
                if (daily.Count >= 30)
                {
                    IndicatorEngine.CalculateAll(daily);
                    var dLast = daily.Last();
                    signal.Daily_RSI = dLast.RSI;
                    signal.Daily_MACD_Hist = dLast.MACD_Hist;
                    signal.Daily_Pattern = dLast.Pattern;
                    signal.Daily_Trend = dLast.EMA_50 > 0
                        ? (dLast.Close > dLast.EMA_50 ? "多頭" : "空頭")
                        : "-";
                }

                // 小時線
                if (hourly.Count >= 30)
                {
                    IndicatorEngine.CalculateAll(hourly);
                    var hLast = hourly.Last();
                    signal.Hourly_RSI = hLast.RSI;
                    signal.Hourly_MACD_Hist = hLast.MACD_Hist;
                    signal.Hourly_Pattern = hLast.Pattern;
                    signal.Hourly_Trend = hLast.EMA_50 > 0
                        ? (hLast.Close > hLast.EMA_50 ? "多頭" : "空頭")
                        : "-";
                }

                // 計算共振分數 (每個框架：多頭+1, 空頭-1, 混合0)
                int score = 0;
                var lines = new List<string>();

                score += ScoreTimeframe(signal.Weekly_Trend, signal.Weekly_RSI, signal.Weekly_MACD_Hist, "週線", lines);
                score += ScoreTimeframe(signal.Daily_Trend, signal.Daily_RSI, signal.Daily_MACD_Hist, "日線", lines);
                score += ScoreTimeframe(signal.Hourly_Trend, signal.Hourly_RSI, signal.Hourly_MACD_Hist, "小時線", lines);

                signal.AlignmentScore = score;
                string scoreEmoji = score >= 2 ? "🟢" : (score <= -2 ? "🔴" : "🟡");
                signal.AlignmentSummary = $"{scoreEmoji} 多空共振分數: {score:+0;-0;0}\n" + string.Join("\n", lines);
            }
            catch (Exception ex) { AppLogger.Log("MultiTimeframeEngine.AnalyzeAsync 失敗", ex); }
            return signal;
        }

        private static int ScoreTimeframe(string trend, double rsi, double macdHist, string label, List<string> lines)
        {
            if (trend == "-") return 0;
            bool bullish = trend == "多頭" && rsi < 70 && macdHist > 0;
            bool bearish = trend == "空頭" && rsi > 30 && macdHist < 0;
            if (bullish) { lines.Add($"✅ {label}: 多頭排列 RSI={rsi:F1} MACD柱={macdHist:F3}"); return 1; }
            if (bearish) { lines.Add($"❌ {label}: 空頭排列 RSI={rsi:F1} MACD柱={macdHist:F3}"); return -1; }
            lines.Add($"⬜ {label}: 中性 RSI={rsi:F1} MACD柱={macdHist:F3}"); return 0;
        }

        /// <summary>將多時間框架信號轉換為 AI 提示詞補充</summary>
        public static string ToPromptContext(MultiTimeframeSignal sig)
        {
            if (sig == null) return "";
            return $"\n【多時間框架共振分析】\n" +
                   $"週線: {sig.Weekly_Trend} | RSI={sig.Weekly_RSI:F1} | MACD柱={sig.Weekly_MACD_Hist:F3} | 形態={sig.Weekly_Pattern}\n" +
                   $"日線: {sig.Daily_Trend} | RSI={sig.Daily_RSI:F1} | MACD柱={sig.Daily_MACD_Hist:F3} | 形態={sig.Daily_Pattern}\n" +
                   $"小時線: {sig.Hourly_Trend} | RSI={sig.Hourly_RSI:F1} | MACD柱={sig.Hourly_MACD_Hist:F3} | 形態={sig.Hourly_Pattern}\n" +
                   $"{sig.AlignmentSummary}\n" +
                   $"👉 請在分析中特別考量多時間框架共振程度，順勢操作時以共振方向為主。\n";
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  回測引擎 (含 Sharpe / Sortino / Kelly)
    // ────────────────────────────────────────────────────────────────────────────
    public static class BacktestEngine
    {
        private const double BuyFee = 0.001425;
        private const double SellFee = 0.001425 + 0.003;
        private const double RiskFreeAnnual = 0.02; // 2% 年化無風險利率

        public class Result
        {
            public double TotalReturn { get; set; }
            public double MaxDrawdown { get; set; }
            public double WinRate { get; set; }
            public int TradeCount { get; set; }
            public double SharpeRatio { get; set; }
            public double SortinoRatio { get; set; }
            public double KellyFraction { get; set; }  // 建議投入比例
            public double AvgWinPct { get; set; }
            public double AvgLossPct { get; set; }
        }

        public static Result RunBacktest(List<MarketData> data)
        {
            double cap = 100000, initial = 100000;
            double pos = 0, entry = 0, maxEq = 100000, mdd = 0;
            int trades = 0, wins = 0;
            var dailyReturns = new List<double>();
            double prevEq = initial;
            var winPcts = new List<double>();
            var lossPcts = new List<double>();

            foreach (var d in data)
            {
                if (d.AgentAction == "Buy" && pos == 0 && d.Close > 0)
                {
                    double invest = cap * 0.95;
                    double costPerShare = d.Close * (1 + BuyFee);
                    double lots = Math.Floor(invest / (costPerShare * 1000));
                    if (lots >= 1) { pos = lots * 1000; cap -= pos * costPerShare; entry = d.Close; trades++; }
                }
                else if (d.AgentAction == "Sell" && pos > 0)
                {
                    double proceeds = pos * d.Close * (1 - SellFee);
                    double pct = (d.Close - entry) / entry;
                    if (d.Close > entry) { wins++; winPcts.Add(pct); } else { lossPcts.Add(pct); }
                    cap += proceeds; pos = 0;
                }

                double eq = cap + pos * d.Close;
                if (prevEq > 0) dailyReturns.Add((eq - prevEq) / prevEq);
                prevEq = eq;
                maxEq = Math.Max(maxEq, eq);
                if (maxEq > 0) mdd = Math.Max(mdd, (maxEq - eq) / maxEq);
            }

            double finalEq = cap + pos * (data.Count > 0 ? data.Last().Close : 0);
            double rfDaily = RiskFreeAnnual / 252.0;

            // Sharpe Ratio
            double sharpe = 0;
            if (dailyReturns.Count > 1)
            {
                double meanR = dailyReturns.Average();
                double stdR = Math.Sqrt(dailyReturns.Select(r => Math.Pow(r - meanR, 2)).Average());
                sharpe = stdR > 0 ? (meanR - rfDaily) / stdR * Math.Sqrt(252) : 0;
            }

            // Sortino Ratio (只用負報酬計算下行標準差)
            double sortino = 0;
            if (dailyReturns.Count > 1)
            {
                double meanR = dailyReturns.Average();
                var downside = dailyReturns.Where(r => r < rfDaily).ToList();
                if (downside.Count > 1)
                {
                    double downStd = Math.Sqrt(downside.Select(r => Math.Pow(r - rfDaily, 2)).Average());
                    sortino = downStd > 0 ? (meanR - rfDaily) / downStd * Math.Sqrt(252) : 0;
                }
            }

            // Kelly Fraction: f* = W - (1-W)/R
            double kelly = 0;
            if (trades > 0)
            {
                double w = winPcts.Count > 0 ? winPcts.Average() : 0;
                double l = lossPcts.Count > 0 ? Math.Abs(lossPcts.Average()) : 0;
                double winRate = (double)wins / trades;
                kelly = l > 0 ? winRate - (1 - winRate) / (w / l) : winRate;
                kelly = Math.Max(0, Math.Min(0.5, kelly)); // 限制在 0~50%
            }

            return new Result
            {
                TotalReturn = (finalEq - initial) / initial,
                MaxDrawdown = mdd,
                WinRate = trades > 0 ? (double)wins / trades : 0,
                TradeCount = trades,
                SharpeRatio = sharpe,
                SortinoRatio = sortino,
                KellyFraction = kelly,
                AvgWinPct = winPcts.Count > 0 ? winPcts.Average() : 0,
                AvgLossPct = lossPcts.Count > 0 ? Math.Abs(lossPcts.Average()) : 0
            };
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  股票篩選引擎
    // ────────────────────────────────────────────────────────────────────────────
    public static class ScreenerEngine
    {
        public static async Task<List<ScreenerResult>> ScanAsync(
            IEnumerable<string> tickers,
            ScreenerCriteria criteria,
            Action<string> onProgress = null)
        {
            var results = new List<ScreenerResult>();
            foreach (var ticker in tickers)
            {
                try
                {
                    onProgress?.Invoke($"掃描中: {ticker}...");
                    var data = await YahooDataService.FetchYahoo(ticker.Trim().ToUpper(), "1d", "3mo");
                    if (data.Count < 30) continue;

                    IndicatorEngine.CalculateAll(data);
                    var last = data.Last();
                    bool isUS = !ticker.EndsWith(".TW") && !ticker.EndsWith(".TWO");
                    string name = await MarketInfoService.GetCompanyName(ticker.Trim().ToUpper(), isUS);

                    // 計算5日均量比
                    double avgVol5 = data.TakeLast(5).Average(x => x.Volume);
                    double avgVol20 = data.TakeLast(20).Average(x => x.Volume);
                    double volRatio = avgVol20 > 0 ? avgVol5 / avgVol20 : 0;

                    // 黃金交叉判斷 (近3根是否出現 MACD 從負轉正)
                    bool macdCrossUp = false;
                    if (data.Count >= 4)
                    {
                        var prev = data[data.Count - 2];
                        macdCrossUp = prev.MACD_Hist < 0 && last.MACD_Hist > 0;
                    }

                    var matched = new List<string>();

                    if (last.RSI >= criteria.RSI_Min && last.RSI <= criteria.RSI_Max)
                        matched.Add($"RSI={last.RSI:F1} ∈ [{criteria.RSI_Min},{criteria.RSI_Max}]");
                    else if (criteria.RSI_Min > 0 || criteria.RSI_Max < 100) continue; // RSI 不符合，跳過

                    if (criteria.MACD_Positive && last.MACD_Hist > 0) matched.Add("MACD柱>0");
                    else if (criteria.MACD_Positive) continue;

                    if (criteria.MACD_CrossUp && macdCrossUp) matched.Add("MACD黃金交叉");
                    else if (criteria.MACD_CrossUp) continue;

                    if (criteria.Above_EMA50 && last.EMA_50 > 0 && last.Close > last.EMA_50) matched.Add("收盤>EMA50");
                    else if (criteria.Above_EMA50) continue;

                    if (criteria.Above_EMA200 && last.EMA_200 > 0 && last.Close > last.EMA_200) matched.Add("收盤>EMA200");
                    else if (criteria.Above_EMA200) continue;

                    if (criteria.BB_Breakout && last.BB_Upper > 0 && last.Close > last.BB_Upper) matched.Add("突破布林上軌");
                    else if (criteria.BB_Breakout) continue;

                    if (criteria.BB_Oversold && last.BB_Lower > 0 && last.Close < last.BB_Lower) matched.Add("跌破布林下軌");
                    else if (criteria.BB_Oversold) continue;

                    if (criteria.Volume_Min_Ratio > 0 && volRatio >= criteria.Volume_Min_Ratio) matched.Add($"量能爆發x{volRatio:F1}");
                    else if (criteria.Volume_Min_Ratio > 0) continue;

                    string trend = last.EMA_200 > 0
                        ? (last.EMA_50 > last.EMA_200 ? "多頭排列" : "空頭排列")
                        : (last.EMA_50 > 0 ? (last.Close > last.EMA_50 ? "收>EMA50" : "收<EMA50") : "-");

                    results.Add(new ScreenerResult
                    {
                        Ticker = ticker.Trim().ToUpper(),
                        CompanyName = string.IsNullOrEmpty(name) ? ticker : name,
                        Close = last.Close,
                        RSI = last.RSI,
                        MACD_Hist = last.MACD_Hist,
                        EMA50 = last.EMA_50,
                        BB_Width = last.BB_Width,
                        Trend = trend,
                        Pattern = last.Pattern,
                        MatchScore = matched.Count,
                        MatchedRules = matched
                    });
                }
                catch (Exception ex) { AppLogger.Log($"ScreenerEngine.ScanAsync {ticker} 失敗", ex); }
            }

            return results.OrderByDescending(r => r.MatchScore).ToList();
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  AI 代理人引擎
    // ────────────────────────────────────────────────────────────────────────────
    public static class AgentEngine
    {
        public static async Task RunAlphaDebate(
            List<MarketData> data, string newsContext, string key, string prompt,
            Action<string> onLog, MultiTimeframeSignal mtfSignal = null,
            CancellationToken ct = default)
        {
            var target = data.TakeLast(30).ToList();
            var sb = new StringBuilder($"【近期市場新聞情緒與基本面估值】\n{newsContext}");

            if (mtfSignal != null)
                sb.Append(MultiTimeframeEngine.ToPromptContext(mtfSignal));

            sb.Append("\n\n【歷史技術數據】\nDate,Close,RSI,MACD_Hist,BBW,Pattern\n");
            foreach (var d in target)
                sb.AppendLine($"{d.Date:MMdd},{d.Close:F1},{d.RSI:F1},{d.MACD_Hist:F3},{d.BB_Width:F1},{d.Pattern}");

            var payload = new
            {
                model = LlmConfig.CurrentModel,          // ← 使用可切換的模型設定
                messages = new[]
                {
                    new { role = "system", content = prompt + "\n🚨請務必用「繁體中文」回覆。\n輸出 JSON: { \"debate_log\": \"...\", \"results\": [ { \"Date\": \"MMdd\", \"Action\": \"Buy/Sell/Hold\", \"Reasoning\": \"...\" } ] }" },
                    new { role = "user", content = sb.ToString() }
                },
                response_format = new { type = "json_object" }
            };

            ct.ThrowIfCancellationRequested();

            // ← 使用共享 AppHttpClients.Llm，不再 new HttpClient
            var req = new HttpRequestMessage(HttpMethod.Post, LlmConfig.BaseUrl)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", key);

            var resp = await AppHttpClients.Llm.SendAsync(req, ct);
            resp.EnsureSuccessStatusCode();
            string raw = await resp.Content.ReadAsStringAsync();
            var outerRoot = JsonDocument.Parse(raw).RootElement;
            string inner = outerRoot.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString();
            var root = JsonDocument.Parse(inner).RootElement;

            if (root.TryGetProperty("debate_log", out var logProp))
                onLog(logProp.GetString() ?? "");

            if (root.TryGetProperty("results", out var resultsProp))
            {
                var results = JsonSerializer.Deserialize<List<LlmResult>>(resultsProp.GetRawText());
                if (results != null)
                    foreach (var r in results)
                    {
                        var m = target.FirstOrDefault(x => x.Date.ToString("MMdd") == r.Date);
                        if (m != null) { m.AgentAction = r.Action ?? "Hold"; m.AgentReasoning = r.Reasoning ?? ""; }
                    }
            }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  強化學習反饋引擎
    // ────────────────────────────────────────────────────────────────────────────
    public static class FeedbackEngine
    {
        private const int MaxPendingDays = 5;

        public static void EvaluatePastActions(string ticker, List<MarketData> history)
        {
            var records = ExcelFeedbackManager.LoadFeedback(ticker);
            bool updated = false;

            for (int i = 0; i < history.Count - 1; i++)
            {
                var current = history[i];
                if (current.AgentAction != "Buy" && current.AgentAction != "Sell") continue;

                var existing = records.FirstOrDefault(r => r.Date.Date == current.Date.Date);
                if (existing == null)
                {
                    existing = new FeedbackRecord { Date = current.Date, Ticker = ticker, Action = current.AgentAction, EntryPrice = current.Close, RSI = current.RSI, Outcome = "Pending", Lesson = "" };
                    records.Add(existing); updated = true;
                }

                if (existing.Outcome == "Pending" && i + 3 < history.Count)
                {
                    double fp = history[i + 3].Close;
                    double ret = current.Close > 0 ? (fp - current.Close) / current.Close : 0;

                    bool resolved = false;
                    if (current.AgentAction == "Buy")
                    {
                        if (ret < -0.03) { existing.Outcome = "Loss"; existing.Lesson = $"{current.Date:MM/dd} Buy (RSI={current.RSI:F1}) 後3天跌 {ret:P2}。❌ 可能誤判假突破。"; resolved = true; }
                        else if (ret > 0.02) { existing.Outcome = "Win"; existing.Lesson = $"{current.Date:MM/dd} Buy 後3天漲 {ret:P2}。✅"; resolved = true; }
                    }
                    else
                    {
                        if (ret > 0.03) { existing.Outcome = "Loss"; existing.Lesson = $"{current.Date:MM/dd} Sell (RSI={current.RSI:F1}) 後3天漲 {ret:P2}。❌ 多頭洗盤太早出。"; resolved = true; }
                        else if (ret < -0.02) { existing.Outcome = "Win"; existing.Lesson = $"{current.Date:MM/dd} Sell 後3天跌 {ret:P2}。✅"; resolved = true; }
                    }

                    if (!resolved && i + MaxPendingDays < history.Count)
                    {
                        double lp = history[Math.Min(i + MaxPendingDays, history.Count - 1)].Close;
                        double lr = current.Close > 0 ? (lp - current.Close) / current.Close : 0;
                        bool win = current.AgentAction == "Buy" ? lr >= 0 : lr <= 0;
                        existing.Outcome = win ? "Win" : "Loss";
                        existing.Lesson = $"{current.Date:MM/dd} {current.AgentAction} 強制結算 {MaxPendingDays}天後 {lr:P2}。";
                        resolved = true;
                    }
                    if (resolved) updated = true;
                }
            }

            if (updated)
            {
                try { ExcelFeedbackManager.SaveFeedback(ticker, records); }
                catch (Exception ex) { AppLogger.Log("FeedbackEngine.Save 失敗", ex); }
            }
        }

        public static string GetRecentLessons(string ticker)
        {
            try
            {
                var records = ExcelFeedbackManager.LoadFeedback(ticker);
                var losses = records.Where(r => r.Outcome == "Loss").OrderByDescending(r => r.Date).Take(3).ToList();
                if (!losses.Any()) return "";
                var sb = new StringBuilder("【⚠️ AI 強化學習：歷史錯誤反思】\n");
                foreach (var l in losses) sb.AppendLine($"- {l.Lesson}");
                sb.AppendLine("👉 請在本次決策中嚴格避免重蹈覆轍！");
                return sb.ToString();
            }
            catch (Exception ex) { AppLogger.Log("FeedbackEngine.GetRecentLessons 失敗", ex); return ""; }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  交易日誌 Excel 管理
    // ────────────────────────────────────────────────────────────────────────────
    public static class ExcelJournalManager
    {
        private static readonly string FilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TradeJournal.xlsx");
        private const string SheetName = "Journal";

        public static List<TradeJournalEntry> LoadAll()
        {
            var list = new List<TradeJournalEntry>();
            if (!File.Exists(FilePath)) return list;
            try
            {
                using (var wb = new XLWorkbook(FilePath))
                {
                    if (!wb.Worksheets.TryGetWorksheet(SheetName, out var ws)) return list;
                    foreach (var row in ws.RangeUsed().RowsUsed().Skip(1))
                    {
                        if (!int.TryParse(row.Cell(1).GetString(), out int id)) continue;
                        list.Add(new TradeJournalEntry
                        {
                            Id = id,
                            TradeDate = DateTime.TryParse(row.Cell(2).GetString(), out var dt) ? dt : DateTime.Today,
                            Ticker = row.Cell(3).GetString(),
                            Direction = row.Cell(4).GetString(),
                            EntryPrice = row.Cell(5).TryGetValue<double>(out var ep) ? ep : 0,
                            ExitPrice = row.Cell(6).TryGetValue<double>(out var xp) ? xp : 0,
                            Quantity = row.Cell(7).TryGetValue<double>(out var q) ? q : 0,
                            AiSuggestion = row.Cell(8).GetString(),
                            MyDecision = row.Cell(9).GetString(),
                            Notes = row.Cell(10).GetString(),
                            StopLossPrice = row.Cell(11).TryGetValue<double>(out var sl) ? sl : 0,
                            TargetPrice = row.Cell(12).TryGetValue<double>(out var tp) ? tp : 0
                        });
                    }
                }
            }
            catch (Exception ex) { AppLogger.Log("ExcelJournalManager.LoadAll 失敗", ex); }
            return list;
        }

        public static void SaveAll(List<TradeJournalEntry> data)
        {
            try
            {
                using (var wb = File.Exists(FilePath) ? new XLWorkbook(FilePath) : new XLWorkbook())
                {
                    if (!wb.Worksheets.TryGetWorksheet(SheetName, out var ws))
                        ws = wb.Worksheets.Add(SheetName);
                    else
                        ws.Clear();

                    ws.Cell(1, 1).Value = "Id"; ws.Cell(1, 2).Value = "Date";
                    ws.Cell(1, 3).Value = "Ticker"; ws.Cell(1, 4).Value = "Direction";
                    ws.Cell(1, 5).Value = "EntryPrice"; ws.Cell(1, 6).Value = "ExitPrice";
                    ws.Cell(1, 7).Value = "Quantity"; ws.Cell(1, 8).Value = "AiSuggestion";
                    ws.Cell(1, 9).Value = "MyDecision"; ws.Cell(1, 10).Value = "Notes"; ws.Cell(1, 11).Value = "StopLossPrice"; ws.Cell(1, 12).Value = "TargetPrice";

                    int r = 2;
                    foreach (var e in data.OrderBy(x => x.TradeDate))
                    {
                        ws.Cell(r, 1).Value = e.Id;
                        ws.Cell(r, 2).Value = e.TradeDate.ToString("yyyy-MM-dd");
                        ws.Cell(r, 3).Value = e.Ticker;
                        ws.Cell(r, 4).Value = e.Direction;
                        ws.Cell(r, 5).Value = e.EntryPrice;
                        ws.Cell(r, 6).Value = e.ExitPrice;
                        ws.Cell(r, 7).Value = e.Quantity;
                        ws.Cell(r, 8).Value = e.AiSuggestion;
                        ws.Cell(r, 9).Value = e.MyDecision;
                        ws.Cell(r, 10).Value = e.Notes;
                        ws.Cell(r, 11).Value = e.StopLossPrice;
                        ws.Cell(r, 12).Value = e.TargetPrice;
                        r++;
                    }
                    wb.SaveAs(FilePath);
                }
            }
            catch (Exception ex) { AppLogger.Log("ExcelJournalManager.SaveAll 失敗", ex); }
        }

        public static int NextId(List<TradeJournalEntry> list) =>
            list.Count > 0 ? list.Max(x => x.Id) + 1 : 1;
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  Excel 歷史快取
    // ────────────────────────────────────────────────────────────────────────────
    public static class ExcelCacheManager
    {
        private static readonly string FilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MarketDataCache.xlsx");

        public static List<string> GetCachedTickers()
        {
            var list = new List<string>();
            if (!File.Exists(FilePath)) return list;
            try
            {
                using (var fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var wb = new XLWorkbook(fs))
                    foreach (var ws in wb.Worksheets) list.Add(ws.Name);
            }
            catch (Exception ex) { AppLogger.Log("ExcelCacheManager.GetCachedTickers 失敗", ex); }
            return list;
        }

        public static List<MarketData> LoadData(string ticker)
        {
            var list = new List<MarketData>();
            if (!File.Exists(FilePath)) return list;
            try
            {
                using (var wb = new XLWorkbook(FilePath))
                {
                    if (wb.Worksheets.TryGetWorksheet(ticker, out var ws))
                        foreach (var row in ws.RangeUsed().RowsUsed().Skip(1))
                            if (DateTime.TryParse(row.Cell(1).GetString(), out DateTime dt))
                                list.Add(new MarketData
                                {
                                    Date = dt,
                                    Open = row.Cell(2).TryGetValue<double>(out var o) ? o : 0,
                                    High = row.Cell(3).TryGetValue<double>(out var h) ? h : 0,
                                    Low = row.Cell(4).TryGetValue<double>(out var l) ? l : 0,
                                    Close = row.Cell(5).TryGetValue<double>(out var c) ? c : 0,
                                    Volume = row.Cell(6).TryGetValue<double>(out var v) ? v : 0
                                });
                }
            }
            catch (Exception ex) { AppLogger.Log("ExcelCacheManager.LoadData 失敗", ex); }
            return list;
        }

        public static void SaveData(string ticker, List<MarketData> data)
        {
            if (data == null || data.Count == 0) return;
            try
            {
                using (var wb = File.Exists(FilePath) ? new XLWorkbook(FilePath) : new XLWorkbook())
                {
                    if (!wb.Worksheets.TryGetWorksheet(ticker, out var ws)) ws = wb.Worksheets.Add(ticker); else ws.Clear();
                    ws.Cell(1, 1).Value = "Date"; ws.Cell(1, 2).Value = "Open"; ws.Cell(1, 3).Value = "High";
                    ws.Cell(1, 4).Value = "Low"; ws.Cell(1, 5).Value = "Close"; ws.Cell(1, 6).Value = "Volume";
                    int r = 2;
                    foreach (var d in data.OrderBy(x => x.Date))
                    {
                        ws.Cell(r, 1).Value = d.Date.ToString("yyyy-MM-dd"); ws.Cell(r, 2).Value = d.Open;
                        ws.Cell(r, 3).Value = d.High; ws.Cell(r, 4).Value = d.Low;
                        ws.Cell(r, 5).Value = d.Close; ws.Cell(r, 6).Value = d.Volume; r++;
                    }
                    wb.SaveAs(FilePath);
                }
            }
            catch (Exception ex) { AppLogger.Log("ExcelCacheManager.SaveData 失敗", ex); }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  Excel 反饋日誌管理
    // ────────────────────────────────────────────────────────────────────────────
    public static class ExcelFeedbackManager
    {
        private static readonly string FilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FeedbackLog.xlsx");

        public static List<FeedbackRecord> LoadFeedback(string ticker)
        {
            var list = new List<FeedbackRecord>();
            if (!File.Exists(FilePath)) return list;
            try
            {
                using (var wb = new XLWorkbook(FilePath))
                {
                    if (wb.Worksheets.TryGetWorksheet(ticker, out var ws))
                        foreach (var row in ws.RangeUsed().RowsUsed().Skip(1))
                            if (DateTime.TryParse(row.Cell(1).GetString(), out DateTime dt))
                                list.Add(new FeedbackRecord { Date = dt, Ticker = row.Cell(2).GetString(), Action = row.Cell(3).GetString(), EntryPrice = row.Cell(4).TryGetValue<double>(out var p) ? p : 0, RSI = row.Cell(5).TryGetValue<double>(out var rsi) ? rsi : 0, Outcome = row.Cell(6).GetString(), Lesson = row.Cell(7).GetString() });
                }
            }
            catch (Exception ex) { AppLogger.Log("ExcelFeedbackManager.LoadFeedback 失敗", ex); }
            return list;
        }

        public static void SaveFeedback(string ticker, List<FeedbackRecord> data)
        {
            if (data == null || data.Count == 0) return;
            try
            {
                using (var wb = File.Exists(FilePath) ? new XLWorkbook(FilePath) : new XLWorkbook())
                {
                    if (!wb.Worksheets.TryGetWorksheet(ticker, out var ws)) ws = wb.Worksheets.Add(ticker); else ws.Clear();
                    ws.Cell(1, 1).Value = "Date"; ws.Cell(1, 2).Value = "Ticker"; ws.Cell(1, 3).Value = "Action";
                    ws.Cell(1, 4).Value = "EntryPrice"; ws.Cell(1, 5).Value = "RSI"; ws.Cell(1, 6).Value = "Outcome"; ws.Cell(1, 7).Value = "Lesson";
                    int r = 2;
                    foreach (var d in data.OrderBy(x => x.Date))
                    {
                        ws.Cell(r, 1).Value = d.Date.ToString("yyyy-MM-dd"); ws.Cell(r, 2).Value = d.Ticker;
                        ws.Cell(r, 3).Value = d.Action; ws.Cell(r, 4).Value = d.EntryPrice;
                        ws.Cell(r, 5).Value = d.RSI; ws.Cell(r, 6).Value = d.Outcome; ws.Cell(r, 7).Value = d.Lesson; r++;
                    }
                    wb.SaveAs(FilePath);
                }
            }
            catch (Exception ex) { AppLogger.Log("ExcelFeedbackManager.SaveFeedback 失敗", ex); }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  當沖資料快取
    // ────────────────────────────────────────────────────────────────────────────
    public static class ExcelIntradayCacheManager
    {
        private static readonly string FilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "IntradayDataCache.xlsx");

        public static List<MarketData> LoadTodayData(string ticker)
        {
            var list = new List<MarketData>();
            if (!File.Exists(FilePath)) return list;
            try
            {
                using (var wb = new XLWorkbook(FilePath))
                {
                    if (wb.Worksheets.TryGetWorksheet(ticker, out var ws))
                        foreach (var row in ws.RangeUsed().RowsUsed().Skip(1))
                            if (DateTime.TryParse(row.Cell(1).GetString(), out DateTime dt) && dt.Date == DateTime.Today)
                                list.Add(new MarketData { Date = dt, Open = row.Cell(2).TryGetValue<double>(out var o) ? o : 0, High = row.Cell(3).TryGetValue<double>(out var h) ? h : 0, Low = row.Cell(4).TryGetValue<double>(out var l) ? l : 0, Close = row.Cell(5).TryGetValue<double>(out var c) ? c : 0, Volume = row.Cell(6).TryGetValue<double>(out var v) ? v : 0 });
                }
            }
            catch (Exception ex) { AppLogger.Log("ExcelIntradayCacheManager.LoadTodayData 失敗", ex); }
            return list;
        }

        public static void SaveData(string ticker, List<MarketData> data)
        {
            if (data == null || data.Count == 0) return;
            try
            {
                using (var wb = File.Exists(FilePath) ? new XLWorkbook(FilePath) : new XLWorkbook())
                {
                    if (!wb.Worksheets.TryGetWorksheet(ticker, out var ws)) ws = wb.Worksheets.Add(ticker); else ws.Clear();
                    ws.Cell(1, 1).Value = "Time"; ws.Cell(1, 2).Value = "Open"; ws.Cell(1, 3).Value = "High";
                    ws.Cell(1, 4).Value = "Low"; ws.Cell(1, 5).Value = "Close"; ws.Cell(1, 6).Value = "Volume";
                    int r = 2;
                    foreach (var d in data.OrderBy(x => x.Date))
                    {
                        ws.Cell(r, 1).Value = d.Date.ToString("yyyy-MM-dd HH:mm:ss"); ws.Cell(r, 2).Value = d.Open;
                        ws.Cell(r, 3).Value = d.High; ws.Cell(r, 4).Value = d.Low;
                        ws.Cell(r, 5).Value = d.Close; ws.Cell(r, 6).Value = d.Volume; r++;
                    }
                    wb.SaveAs(FilePath);
                }
            }
            catch (Exception ex) { AppLogger.Log("ExcelIntradayCacheManager.SaveData 失敗", ex); }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  對話記錄
    // ────────────────────────────────────────────────────────────────────────────

    // ────────────────────────────────────────────────────────────────────────────
    //  停損停利建議引擎（ATR 法 + Fibonacci 法）
    // ────────────────────────────────────────────────────────────────────────────
    public static class StopLossSuggestionEngine
    {
        private const double AtrMultiplierStop = 1.5;
        private const double AtrMultiplierTarget1 = 1.5;
        private const double AtrMultiplierTarget2 = 3.0;
        private const double AtrMultiplierTarget3 = 4.5;

        public static RiskRewardSuggestion CalcByATR(List<MarketData> data, string direction = "Buy")
        {
            if (data == null || data.Count < 15) return null;
            var last = data.Last();
            double atr = last.ATR > 0 ? last.ATR : data.TakeLast(14).Average(d => d.High - d.Low);
            double entry = last.Close;
            double stop, t1, t2, t3;
            if (direction == "Buy")
            {
                stop = entry - atr * AtrMultiplierStop;
                t1 = entry + atr * AtrMultiplierTarget1;
                t2 = entry + atr * AtrMultiplierTarget2;
                t3 = entry + atr * AtrMultiplierTarget3;
            }
            else
            {
                stop = entry + atr * AtrMultiplierStop;
                t1 = entry - atr * AtrMultiplierTarget1;
                t2 = entry - atr * AtrMultiplierTarget2;
                t3 = entry - atr * AtrMultiplierTarget3;
            }
            double riskPct = entry > 0 ? Math.Abs(entry - stop) / entry : 0;
            string dir = direction == "Buy" ? "做多" : "做空";
            return new RiskRewardSuggestion
            {
                EntryPrice = entry,
                Method = "ATR",
                StopLoss = Math.Round(stop, 2),
                Target1 = Math.Round(t1, 2),
                Target2 = Math.Round(t2, 2),
                Target3 = Math.Round(t3, 2),
                AtrValue = Math.Round(atr, 2),
                RiskPct = riskPct,
                Summary = $"【ATR法 · {dir}】進場 {entry:F2}  ATR={atr:F2}\n" +
                           $"🛑 停損: {stop:F2}  ({riskPct:P1} 風險)\n" +
                           $"🎯 T1 (1:1): {t1:F2}\n🎯 T2 (1:2): {t2:F2}\n🎯 T3 (1:3): {t3:F2}"
            };
        }

        public static RiskRewardSuggestion CalcByFibonacci(List<MarketData> data, string direction = "Buy")
        {
            if (data == null || data.Count < 10) return null;
            var fibs = FibonacciEngine.Calculate(data, Math.Min(120, data.Count));
            if (fibs == null || fibs.Count < 3) return CalcByATR(data, direction);
            double entry = data.Last().Close;
            double atr = data.Last().ATR > 0 ? data.Last().ATR : data.TakeLast(14).Average(d => d.High - d.Low);

            var stops = direction == "Buy"
                ? fibs.Where(f => f.Price < entry).OrderByDescending(f => f.Price).ToList()
                : fibs.Where(f => f.Price > entry).OrderBy(f => f.Price).ToList();
            var targets = direction == "Buy"
                ? fibs.Where(f => f.Price > entry).OrderBy(f => f.Price).ToList()
                : fibs.Where(f => f.Price < entry).OrderByDescending(f => f.Price).ToList();

            if (stops.Count == 0 || targets.Count == 0) return CalcByATR(data, direction);
            double stop = stops[0].Price;
            double t1 = targets.Count > 0 ? targets[0].Price : entry + atr * 1.5;
            double t2 = targets.Count > 1 ? targets[1].Price : entry + atr * 3.0;
            double t3 = targets.Count > 2 ? targets[2].Price : entry + atr * 4.5;
            double riskPct = entry > 0 ? Math.Abs(entry - stop) / entry : 0;
            double rr1 = riskPct > 0 ? Math.Abs(t1 - entry) / Math.Abs(entry - stop) : 0;
            string dir = direction == "Buy" ? "做多" : "做空";
            return new RiskRewardSuggestion
            {
                EntryPrice = entry,
                Method = "Fibonacci",
                StopLoss = Math.Round(stop, 2),
                Target1 = Math.Round(t1, 2),
                Target2 = Math.Round(t2, 2),
                Target3 = Math.Round(t3, 2),
                AtrValue = Math.Round(atr, 2),
                RiskPct = riskPct,
                FibStopLabel = stops[0].Label,
                FibTargetLabel = targets.Count > 0 ? targets[0].Label : "",
                Summary = $"【Fibonacci法 · {dir}】進場 {entry:F2}\n" +
                          $"🛑 停損: {stop:F2} ({stops[0].Label}, {riskPct:P1})\n" +
                          $"🎯 T1 ({targets[0].Label}): {t1:F2}  [風報比 1:{rr1:F1}]\n" +
                          $"🎯 T2 ({(targets.Count > 1 ? targets[1].Label : "ext")}): {t2:F2}\n" +
                          $"🎯 T3 ({(targets.Count > 2 ? targets[2].Label : "ext")}): {t3:F2}"
            };
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  止盈止損警報引擎
    // ────────────────────────────────────────────────────────────────────────────
    public static class AlertEngine
    {
        private static List<StopLossAlert> _alerts = new List<StopLossAlert>();
        private static bool _loaded = false;
        private static readonly string FilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Alerts.json");

        public static List<StopLossAlert> GetActiveAlerts()
        { EnsureLoaded(); return _alerts.Where(a => a.IsActive && !a.StopTriggered && !a.TargetTriggered).ToList(); }

        public static void SyncFromJournal(List<TradeJournalEntry> entries)
        {
            EnsureLoaded();
            foreach (var e in entries.Where(e => e.ExitPrice == 0 && (e.StopLossPrice > 0 || e.TargetPrice > 0)))
            {
                if (!_alerts.Any(a => a.JournalId == e.Id))
                    _alerts.Add(new StopLossAlert
                    {
                        JournalId = e.Id,
                        Ticker = e.Ticker,
                        Direction = e.Direction,
                        EntryPrice = e.EntryPrice,
                        StopLossPrice = e.StopLossPrice,
                        TargetPrice = e.TargetPrice,
                        IsActive = true
                    });
                else
                {
                    var a = _alerts.First(a => a.JournalId == e.Id);
                    a.StopLossPrice = e.StopLossPrice; a.TargetPrice = e.TargetPrice; a.IsActive = true;
                }
            }
            var closedIds = entries.Where(e => e.ExitPrice > 0).Select(e => e.Id).ToHashSet();
            foreach (var a in _alerts.Where(a => closedIds.Contains(a.JournalId))) a.IsActive = false;
            Save();
        }

        public static List<StopLossAlert> CheckPrice(string ticker, double currentPrice)
        {
            EnsureLoaded();
            var triggered = new List<StopLossAlert>();
            foreach (var a in _alerts.Where(a => a.IsActive && a.Ticker == ticker && !a.StopTriggered && !a.TargetTriggered))
            {
                bool hitStop = a.StopLossPrice > 0 && (a.Direction == "Buy" ? currentPrice <= a.StopLossPrice : currentPrice >= a.StopLossPrice);
                bool hitTarget = a.TargetPrice > 0 && (a.Direction == "Buy" ? currentPrice >= a.TargetPrice : currentPrice <= a.TargetPrice);
                if (hitStop) { a.StopTriggered = true; a.TriggeredAt = DateTime.Now; a.TriggeredType = "StopLoss"; triggered.Add(a); }
                if (hitTarget) { a.TargetTriggered = true; a.TriggeredAt = DateTime.Now; a.TriggeredType = "Target"; triggered.Add(a); }
            }
            if (triggered.Count > 0) Save();
            return triggered;
        }

        public static void DismissAlert(int journalId)
        {
            var a = _alerts.FirstOrDefault(x => x.JournalId == journalId);
            if (a != null) { a.IsActive = false; Save(); }
        }

        // 公開的 SaveAlerts（供 SentimentAwareTradingEngine.ApplyVixJumpStop 呼叫後使用）
        public static void SaveAlerts(List<StopLossAlert> updated)
        {
            EnsureLoaded();
            foreach (var upd in updated)
            {
                var existing = _alerts.FirstOrDefault(a => a.JournalId == upd.JournalId);
                if (existing != null)
                {
                    existing.StopLossPrice = upd.StopLossPrice;
                    existing.TargetPrice = upd.TargetPrice;
                    existing.IsActive = upd.IsActive;
                }
            }
            Save();
        }

        private static void EnsureLoaded()
        {
            if (_loaded) return; _loaded = true;
            if (!File.Exists(FilePath)) return;
            try { _alerts = JsonSerializer.Deserialize<List<StopLossAlert>>(File.ReadAllText(FilePath)) ?? new List<StopLossAlert>(); }
            catch (Exception ex) { AppLogger.Log("AlertEngine.Load 失敗", ex); }
        }

        private static void Save()
        {
            try { File.WriteAllText(FilePath, JsonSerializer.Serialize(_alerts, new JsonSerializerOptions { WriteIndented = true })); }
            catch (Exception ex) { AppLogger.Log("AlertEngine.Save 失敗", ex); }
        }
    }

    public static class ExcelChatLogger
    {
        private static readonly string FilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ChatHistory.xlsx");
        private static readonly SemaphoreSlim _lock = new SemaphoreSlim(1, 1);

        public static void LogMessage(string ticker, string role, string message)
        {
            Task.Run(async () =>
            {
                await _lock.WaitAsync();
                try
                {
                    using (var wb = File.Exists(FilePath) ? new XLWorkbook(FilePath) : new XLWorkbook())
                    {
                        if (!wb.Worksheets.TryGetWorksheet(ticker, out var ws))
                        { ws = wb.Worksheets.Add(ticker); ws.Cell(1, 1).Value = "Timestamp"; ws.Cell(1, 2).Value = "Role"; ws.Cell(1, 3).Value = "Message"; }
                        int nr = (ws.LastRowUsed()?.RowNumber() ?? 1) + 1;
                        ws.Cell(nr, 1).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        ws.Cell(nr, 2).Value = role; ws.Cell(nr, 3).Value = message;
                        wb.SaveAs(FilePath);
                    }
                }
                catch (Exception ex) { AppLogger.Log("ExcelChatLogger 失敗", ex); }
                finally { _lock.Release(); }
            });
        }
    }
    // ════════════════════════════════════════════════════════════════════════════
    //  ② PositionSizingEngine
    //  Kelly 公式 + 金字塔加碼 + 分批出場 + 跟蹤停損
    // ════════════════════════════════════════════════════════════════════════════
    public static class PositionSizingEngine
    {
        private const double MinLotShares = 1000;    // 台股一張 = 1000 股
        private const double MaxKellyClamp = 0.40;   // Kelly 上限 40%

        // ── Kelly 倉位計算 ────────────────────────────────────────────────────
        /// <summary>
        /// 根據 Kelly Fraction 和停損距離計算建議進場量。
        /// lotSize=1000 用於台股；美股傳 1。
        /// </summary>
        public static PositionSizingResult CalcEntrySize(
            double capital,
            double kellyFraction,
            double entryPrice,
            double stopLoss,
            double lotSize = 1000)
        {
            double kelly = Math.Max(0, Math.Min(MaxKellyClamp, kellyFraction));
            double riskPerShare = Math.Abs(entryPrice - stopLoss);
            double stopLossPct = entryPrice > 0 ? riskPerShare / entryPrice : 0;

            // 風險金額 = 資金 × Kelly
            double riskAmount = capital * kelly;

            // 建議股數：風險金額 / 每股風險，再對齊整張
            double rawQty = stopLossPct > 0 && entryPrice > 0
                ? riskAmount / riskPerShare
                : capital * kelly / entryPrice;

            double qty = lotSize > 1
                ? Math.Floor(rawQty / lotSize) * lotSize
                : Math.Floor(rawQty);

            string summary;
            if (qty < lotSize && lotSize > 1)
            {
                // 資金不足一張時給出最低限建議
                qty = lotSize;
                summary = $"⚠️ 依 Kelly={kelly:P0} 計算股數不足一張，建議最低 1 張試倉";
            }
            else
            {
                summary = $"Kelly={kelly:P0}  停損幅度={stopLossPct:P2}  " +
                          $"建議進場 {qty:N0} 股（市值 ${qty * entryPrice:N0}）  " +
                          $"最大風險 ${riskAmount:N0}";
            }

            return new PositionSizingResult
            {
                Capital = capital,
                KellyFraction = kelly,
                RiskAmount = riskAmount,
                EntryPrice = entryPrice,
                StopLossPrice = stopLoss,
                StopLossPct = stopLossPct,
                RecommendedQty = qty,
                MaxPositionValue = qty * entryPrice,
                Summary = summary
            };
        }

        // ── 金字塔加碼：三批進場策略 ─────────────────────────────────────────
        /// <summary>
        /// 生成分批進場計畫：突破確認→加碼→最後確認三個批次
        /// 預設分配：50% / 30% / 20%
        /// </summary>
        public static List<PyramidLevel> GeneratePyramidEntry(
            double basePrice,
            double totalQty,
            int levels = 3,
            bool isLong = true,
            double atr = 0)
        {
            var result = new List<PyramidLevel>();
            // 各層分配比例（合計 = 1）
            var allocs = levels == 2 ? new[] { 0.60, 0.40 }
                       : levels == 3 ? new[] { 0.50, 0.30, 0.20 }
                       : new[] { 0.40, 0.30, 0.20, 0.10 };

            double step = atr > 0 ? atr * 0.3 : basePrice * 0.015;   // 預設每批間距 1.5%

            for (int i = 0; i < Math.Min(levels, allocs.Length); i++)
            {
                double price = isLong
                    ? basePrice + step * i          // 多頭：越漲越加碼（確認突破）
                    : basePrice - step * i;          // 空頭：越跌越加碼

                double qty = Math.Floor(totalQty * allocs[i]);
                string label = i == 0
                    ? $"第①批：{allocs[i]:P0} 試倉（突破確認）"
                    : i == levels - 1
                        ? $"第{(char)('①' + i)}批：{allocs[i]:P0} 最後確認加碼"
                        : $"第{(char)('①' + i)}批：{allocs[i]:P0} 動能加碼";

                result.Add(new PyramidLevel
                {
                    Batch = i + 1,
                    Price = price,
                    QtyPct = allocs[i],
                    Qty = qty,
                    Label = label
                });
            }
            return result;
        }

        // ── 分批出場計畫（分層獲利了結）─────────────────────────────────────
        /// <summary>
        /// 根據進場價和停損，生成 1R/2R/3R 三層出場計畫
        /// </summary>
        public static List<ExitLevel> GenerateLayeredExit(
            double entryPrice,
            double stopLoss,
            double totalQty,
            bool isLong = true)
        {
            double risk = Math.Abs(entryPrice - stopLoss);
            if (risk <= 0) return new List<ExitLevel>();

            var layers = new[]
            {
                (Layer: 1, R: 1.0, Pct: 0.40, Rationale: "1R 第一目標：快速鎖利 40%"),
                (Layer: 2, R: 2.0, Pct: 0.35, Rationale: "2R 第二目標：主要獲利區 35%"),
                (Layer: 3, R: 3.0, Pct: 0.25, Rationale: "3R 第三目標：讓利奔跑 25%")
            };

            return layers.Select(l => new ExitLevel
            {
                Layer = l.Layer,
                Price = isLong
                    ? entryPrice + risk * l.R
                    : entryPrice - risk * l.R,
                QtyPct = l.Pct,
                Rationale = l.Rationale + $"（出場 {totalQty * l.Pct:N0} 股）"
            }).ToList();
        }

        // ── 跟蹤停損計算 ─────────────────────────────────────────────────────
        /// <summary>
        /// 基於 ATR 的跟蹤停損：在最高/最低點往回拉 atrMultiplier × ATR
        /// </summary>
        public static double CalcTrailingStop(
            List<MarketData> data,
            double entryPrice,
            double atrMultiplier = 2.0,
            bool isLong = true)
        {
            if (data == null || data.Count == 0) return entryPrice * (isLong ? 0.95 : 1.05);

            var relevant = data.Where(d => d.Date >= data.Last().Date.AddDays(-60)).ToList();
            double atr = relevant.Count > 0 ? relevant.Average(d => d.ATR) : data.Last().ATR;
            if (atr <= 0) atr = data.Last().Close * 0.02;

            if (isLong)
            {
                double highestClose = relevant.Max(d => d.High);
                return Math.Max(entryPrice * 0.95, highestClose - atrMultiplier * atr);
            }
            else
            {
                double lowestClose = relevant.Min(d => d.Low);
                return Math.Min(entryPrice * 1.05, lowestClose + atrMultiplier * atr);
            }
        }

        // ── 倉位計畫摘要（給 AI Prompt 用）──────────────────────────────────
        public static string ToPromptContext(
            PositionSizingResult sizing,
            List<PyramidLevel> pyramid,
            List<ExitLevel> exits,
            double trailingStop)
        {
            var sb = new StringBuilder();
            sb.AppendLine("【倉位管理建議】");
            sb.AppendLine(sizing.Summary);

            if (pyramid?.Count > 0)
            {
                sb.AppendLine($"▶ 分批進場計畫（金字塔加碼）：");
                foreach (var p in pyramid)
                    sb.AppendLine($"  {p.Label}  目標價 {p.Price:F2}  股數 {p.Qty:N0}");
            }

            if (exits?.Count > 0)
            {
                sb.AppendLine($"◀ 分層獲利了結計畫：");
                foreach (var x in exits)
                    sb.AppendLine($"  {x.Rationale}  出場價 {x.Price:F2}");
            }

            if (trailingStop > 0)
                sb.AppendLine($"🛡 跟蹤停損觸發線：{ trailingStop: F2} ");

            sb.AppendLine("👉 請根據以上倉位計畫，在 Buy/Sell/Hold 之外補充「建議倉位比例」和「加碼時機」。");
            return sb.ToString();
        }

        // ── ATR 計算輔助（供外部調用）────────────────────────────────────────
        public static double CalcATR(List<MarketData> data, int period = 14)
        {
            if (data == null || data.Count < 2) return 0;
            var recent = data.TakeLast(period + 1).ToList();
            double sum = 0; int count = 0;
            for (int i = 1; i < recent.Count; i++)
            {
                double tr = Math.Max(recent[i].High - recent[i].Low,
                            Math.Max(Math.Abs(recent[i].High - recent[i - 1].Close),
                                     Math.Abs(recent[i].Low - recent[i - 1].Close)));
                sum += tr; count++;
            }
            return count > 0 ? sum / count : 0;
        }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  ③ SentimentAwareTradingEngine
    //  情緒感知風控：VIX → 倉位上限 + Prompt 強化 + VIX 跳升停損
    // ════════════════════════════════════════════════════════════════════════════
    public static class SentimentAwareTradingEngine
    {
        // ── 情緒 → 風控上限 ───────────────────────────────────────────────────
        /// <summary>
        /// 根據 VIX 和 Fear/Greed 計算最大持倉比例和最大容忍回撤
        /// </summary>
        public static SentimentRiskLimit CalcSentimentRiskLimit(
            MarketSentiment sentiment,
            double vixPrev5dAvg = 0)
        {
            if (sentiment == null)
                return new SentimentRiskLimit { MaxPositionPct = 0.80, MaxDrawdownPct = 0.10, Regime = "無資料" };

            double vix = sentiment.VIX;
            double fg = sentiment.FearGreedScore;

            // VIX 加速度（今日 vs 5日均）
            double vixAccel = vixPrev5dAvg > 0 ? (vix - vixPrev5dAvg) / vixPrev5dAvg : 0;
            bool vixJump = vixAccel > 0.20;    // VIX 急升 20% 以上

            string regime;
            double maxPos, maxDD;

            if (vix > 40 || fg < 15)            // 極度恐慌
            {
                regime = "⚠️ 極度恐慌"; maxPos = 0.20; maxDD = 0.06;
            }
            else if (vix > 30 || fg < 25)       // 恐慌
            {
                regime = "🔴 恐慌"; maxPos = 0.35; maxDD = 0.08;
            }
            else if (vix > 20 || fg < 40)       // 偏謹慎
            {
                regime = "🟡 偏謹慎"; maxPos = 0.55; maxDD = 0.10;
            }
            else if (vix < 12 && fg > 75)       // 極度貪婪（也要控制）
            {
                regime = "🔥 極度貪婪"; maxPos = 0.60; maxDD = 0.12;
            }
            else if (fg > 60)                   // 貪婪
            {
                regime = "💚 貪婪"; maxPos = 0.75; maxDD = 0.12;
            }
            else                                 // 中性
            {
                regime = "⬜ 中性"; maxPos = 0.80; maxDD = 0.12;
            }

            // VIX 急跳 → 額外降倉
            if (vixJump) { maxPos *= 0.60; regime += " + VIX急升"; }

            // 建構 Prompt 修飾語
            var prompt = BuildPromptModifier(regime, maxPos, maxDD, vixJump, vixAccel, sentiment);

            return new SentimentRiskLimit
            {
                MaxPositionPct = maxPos,
                MaxDrawdownPct = maxDD,
                Regime = regime,
                PromptModifier = prompt,
                VixJumpAlert = vixJump,
                VixAcceleration = vixAccel
            };
        }

        private static string BuildPromptModifier(
            string regime, double maxPos, double maxDD,
            bool vixJump, double vixAccel, MarketSentiment s)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"【🔔 市場情緒風控指令（優先級最高）】");
            sb.AppendLine($"當前情緒政策：{regime}");
            sb.AppendLine($"▶ 本次分析最大持倉限制：{maxPos:P0}（超過請強制降至此水位）");
            sb.AppendLine($"▶ 可容忍最大回撤：{maxDD:P0}（超過立刻止損，不等待反彈）");

            if (vixJump)
            {
                sb.AppendLine($"🚨 VIX 急升警報：VIX 加速度 {vixAccel:+0.0%;-0.0%}！");
                sb.AppendLine($"   → 所有持倉須立即設定不低於 {maxDD * 0.5:P0} 的緊急停損");
                sb.AppendLine($"   → 禁止在 VIX 急升期間加碼，等待 VIX 回穩再評估");
            }

            if (s.FearGreedScore < 20)
                sb.AppendLine($"💡 極度恐懼可能是買點，但須等恐懼緩和（FG > 25）後再進場，不搶刀。");
            else if (s.FearGreedScore > 80)
                sb.AppendLine($"💡 極度貪婪市場，降低新進場積極程度，優先縮短持倉時間。");

            sb.AppendLine($"👉 請在回覆中明確說明「是否符合 {maxPos:P0} 持倉上限」。");
            return sb.ToString();
        }

        // ── VIX 跳升 → 收緊現有停損警報 ─────────────────────────────────────
        /// <summary>
        /// 當 VIX 相對前一日大幅跳升時，收緊所有活躍警報的停損到 ATR × 1.5
        /// </summary>
        public static int ApplyVixJumpStop(
            List<StopLossAlert> activeAlerts,
            List<MarketData> histData,
            double vixCurrent,
            double vixPrev)
        {
            if (activeAlerts == null || activeAlerts.Count == 0) return 0;
            double vixChange = vixPrev > 0 ? (vixCurrent - vixPrev) / vixPrev : 0;
            if (vixChange < 0.15) return 0;   // 15% 以下不觸發

            double atr = histData?.Count > 5
                ? histData.TakeLast(14).Average(d => d.ATR)
                : 0;
            int updated = 0;

            foreach (var alert in activeAlerts.Where(a => a.IsActive))
            {
                // 收緊：新停損 = 進場價往不利方向移 1.5 × ATR（最多移到原始停損）
                double tighter;
                if (alert.Direction == "Buy")
                {
                    tighter = atr > 0
                        ? alert.EntryPrice - atr * 1.5
                        : alert.EntryPrice * 0.97;
                    // 只收緊（不放寬），且不超過原有停損的有利方向
                    if (tighter > alert.StopLossPrice)
                    {
                        alert.StopLossPrice = tighter;
                        updated++;
                    }
                }
                else
                {
                    tighter = atr > 0
                        ? alert.EntryPrice + atr * 1.5
                        : alert.EntryPrice * 1.03;
                    if (tighter < alert.StopLossPrice)
                    {
                        alert.StopLossPrice = tighter;
                        updated++;
                    }
                }
            }
            return updated;
        }

        // ── Short Squeeze 預警（Fear→Greed 轉折偵測）─────────────────────────
        /// <summary>
        /// 比較最近兩次情緒讀值，偵測 Fear→Greed 反轉
        /// 返回 (是否即將反轉, 信心分數 0~100, 說明)
        /// </summary>
        public static (bool Signal, int Confidence, string Message) DetectReversalWarning(
            MarketSentiment prev, MarketSentiment curr)
        {
            if (prev == null || curr == null) return (false, 0, "");

            double fgChange = curr.FearGreedScore - prev.FearGreedScore;
            double vixChange = prev.VIX > 0 ? (curr.VIX - prev.VIX) / prev.VIX : 0;

            bool fearToGreedSignal = prev.FearGreedScore < 30 && fgChange > 8;   // 恐懼區 FG 快速上升
            bool vixReversal = prev.VIX > 25 && vixChange < -0.08;          // VIX 急速下降

            if (!fearToGreedSignal && !vixReversal) return (false, 0, "");

            int confidence = 0;
            var msgs = new List<string>();
            if (fearToGreedSignal) { confidence += 50; msgs.Add($"FG 從 {prev.FearGreedScore:F0} 升至 {curr.FearGreedScore:F0}（+{fgChange:F0}）"); }
            if (vixReversal) { confidence += 50; msgs.Add($"VIX 從 {prev.VIX:F1} 降至 {curr.VIX:F1}（{vixChange:P1}）"); }

            string msg = $"🔄 情緒反轉預警 (信心 {confidence}%)：" + string.Join("  ", msgs) +
                         "  → 空頭回補（Short Squeeze）風險升高！";
            return (true, confidence, msg);
        }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  ④ SectorCorrelationEngine
    //  板塊相關係數矩陣 + 領先/滯後偵測 + 相對強度排名
    //  輸入：各板塊近期歷史資料（由 SectorRotationService 的快照補充）
    // ════════════════════════════════════════════════════════════════════════════
    public static class SectorCorrelationEngine
    {
        // ── 相關係數矩陣 ─────────────────────────────────────────────────────
        /// <summary>
        /// 計算 n 個板塊之間的日報酬相關係數矩陣（Pearson）。
        /// sectorReturns[i] = 板塊 i 的日報酬序列（已對齊日期）
        /// </summary>
        public static double[,] CalcCorrelationMatrix(List<double[]> sectorReturns)
        {
            int n = sectorReturns.Count;
            var mat = new double[n, n];
            for (int i = 0; i < n; i++)
                for (int j = 0; j < n; j++)
                    mat[i, j] = i == j ? 1.0 : PearsonCorrelation(sectorReturns[i], sectorReturns[j]);
            return mat;
        }

        private static double PearsonCorrelation(double[] x, double[] y)
        {
            int len = Math.Min(x.Length, y.Length);
            if (len < 3) return 0;
            double meanX = x.Take(len).Average();
            double meanY = y.Take(len).Average();
            double cov = 0, varX = 0, varY = 0;
            for (int i = 0; i < len; i++)
            {
                double dx = x[i] - meanX, dy = y[i] - meanY;
                cov += dx * dy; varX += dx * dx; varY += dy * dy;
            }
            double denom = Math.Sqrt(varX * varY);
            return denom > 0 ? cov / denom : 0;
        }

        // ── 板塊相對強度（RS）排名 ────────────────────────────────────────────
        /// <summary>
        /// RS = 板塊 20 日報酬 / SPY 20 日報酬。
        /// snapshots：(ticker, name, emoji, changePct, price) 的快照
        /// spyReturn：SPY 同期報酬（傳 0 時降級為純漲幅排名）
        /// </summary>
        public static SectorCorrelationResult BuildResult(
            List<(string Ticker, string Name, string Emoji, double Return20d, double Price)> sectors,
            double spyReturn20d,
            List<double[]> returnSeries = null)
        {
            var result = new SectorCorrelationResult();

            // 相對強度排名
            var ranking = sectors.Select(s =>
            {
                double rs = spyReturn20d != 0 ? s.Return20d / Math.Abs(spyReturn20d) : s.Return20d;
                return (s.Ticker, s.Name , rs, s.Return20d);
            }).OrderByDescending(x => x.rs).ToList();
            result.Ranking = ranking;
            result.SectorNames = sectors.Select(s => s.Name).ToArray();

            // 領先/滯後
            result.LeadingSector = ranking.FirstOrDefault().Name ?? "";
            result.LaggingSector = ranking.LastOrDefault().Name ?? "";

            // 相關係數矩陣（若有資料序列）
            if (returnSeries?.Count >= 2)
            {
                result.Matrix = CalcCorrelationMatrix(returnSeries);

                // 輪動信號：找強弱差距大且最近反轉的板塊對
                var topRS = ranking.Take(3).Select(r => r.Ticker).ToHashSet();
                var bottomRS = ranking.TakeLast(3).Select(r => r.Ticker).ToHashSet();
                var signals = new List<(string, string, double)>();

                foreach (var top in ranking.Take(2))
                    foreach (var bot in ranking.TakeLast(2))
                    {
                        double gap = top.rs - bot.rs;
                        if (gap > 0.15)   // 強弱差距 > 15%
                        {
                            double conf = Math.Min(1.0, gap * 3);   // 最高 100%
                            signals.Add((bot.Name, top.Name, conf));
                        }
                    }
                result.RotationSignals = signals;
            }

            return result;
        }

        // ── 生成 AI Prompt 板塊情境 ──────────────────────────────────────────
        public static string ToPromptContext(SectorCorrelationResult r)
        {
            if (r == null || r.Ranking == null || r.Ranking.Count == 0) return "";
            var sb = new StringBuilder();
            sb.AppendLine("【板塊輪動情境】");
            sb.Append("強勢板塊：");
            sb.AppendLine(string.Join("  ", r.Ranking.Take(3).Select(x => $"{x.Name}(RS={x.RS:+0.00;-0.00})")));
            sb.Append("弱勢板塊：");
            sb.AppendLine(string.Join("  ", r.Ranking.TakeLast(3).Reverse().Select(x => $"{x.Name}(RS={x.RS:+0.00;-0.00})")));
            if (r.RotationSignals?.Count > 0)
            {
                sb.AppendLine("輪動訊號：");
                foreach (var (from, to, conf) in r.RotationSignals.Take(2))
                    sb.AppendLine($"  資金可能從 {from} → {to}（信心 {conf:P0}）");
            }
            return sb.ToString();
        }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  ⑤ VolatilityAdaptiveEngine
    //  VIX 動態調整 ATR + 波動率百分位 + 衝擊成本估算
    // ════════════════════════════════════════════════════════════════════════════
    public static class VolatilityAdaptiveEngine
    {
        // ── VIX 調整 ATR ─────────────────────────────────────────────────────
        /// <summary>
        /// 標準化 ATR 再依 VIX 相對歷史均值進行縮放。
        /// ATR_adaptive = ATR_raw × sqrt(VIX_current / VIX_histAvg)
        /// </summary>
        public static double GetAdaptiveAtr(
            List<MarketData> data,
            double vixCurrent,
            double vixHistAvg = 20.0,
            int period = 14)
        {
            if (data == null || data.Count < period) return 0;
            double rawAtr = data.TakeLast(period).Average(d => d.ATR);
            if (rawAtr <= 0) return 0;
            double vixRatio = vixHistAvg > 0 ? vixCurrent / vixHistAvg : 1.0;
            return rawAtr * Math.Sqrt(vixRatio);
        }

        // ── 波動率百分位 ─────────────────────────────────────────────────────
        /// <summary>
        /// 計算當前 ATR 在歷史分布中的百分位（0~100）。
        /// 100 = 歷史最高波動，0 = 最低。
        /// </summary>
        public static double CalcVolatilityPercentile(List<MarketData> data, int period = 14)
        {
            if (data == null || data.Count < period * 3) return 50.0;

            var atrHistory = data
                .Select(d => d.ATR)
                .Where(a => a > 0)
                .ToList();

            if (atrHistory.Count < 10) return 50.0;

            double currentAtr = atrHistory.TakeLast(period).Average();
            int below = atrHistory.Count(a => a <= currentAtr);
            return (double)below / atrHistory.Count * 100.0;
        }

        // ── 建議停損倍數（依波動率百分位）────────────────────────────────────
        /// <summary>
        /// 低波動 → 小停損倍數；高波動 → 大停損倍數，避免頻繁洗出
        /// </summary>
        public static double GetSuggestedAtrMultiplier(double volPercentile)
        {
            if (volPercentile >= 90) return 3.5;      // 極高波動
            if (volPercentile >= 75) return 3.0;      // 高波動
            if (volPercentile >= 50) return 2.5;      // 正常偏高
            if (volPercentile >= 25) return 2.0;      // 正常偏低
            return 1.5;                                // 低波動
        }

        // ── 衝擊成本估算 ─────────────────────────────────────────────────────
        /// <summary>
        /// 小股票 / 大單下單時的市場衝擊成本估算。
        /// Kyle's Lambda 簡化版：impact ≈ σ × sqrt(orderSize/dailyVolume)
        /// 回傳：預估衝擊成本（以股價 % 表示）
        /// </summary>
        public static double EstimateImpactCost(
            double orderSize,
            double dailyVolume,
            double volatility)
        {
            if (dailyVolume <= 0 || volatility <= 0) return 0;
            return volatility * Math.Sqrt(orderSize / dailyVolume);
        }

        // ── 完整波動率剖面 ────────────────────────────────────────────────────
        public static VolatilityProfile BuildProfile(
            List<MarketData> data,
            double vixCurrent,
            double vixHistAvg = 20.0,
            double entryPrice = 0)
        {
            double rawAtr = data?.Count > 14 ? data.TakeLast(14).Average(d => d.ATR) : 0;
            double adaptAtr = GetAdaptiveAtr(data, vixCurrent, vixHistAvg);
            double pct = CalcVolatilityPercentile(data);
            double mult = GetSuggestedAtrMultiplier(pct);

            string regime = pct >= 90 ? "🔥 極端波動"
                          : pct >= 75 ? "⚠️ 高波動"
                          : pct >= 40 ? "⬜ 正常波動"
                          : "💤 低波動";

            string summary = $"ATR={rawAtr:F2}  自適應ATR={adaptAtr:F2}  " +
                             $"波動率百分位={pct:F0}%  建議停損倍數={mult:F1}x  {regime}";

            if (pct >= 75)
                summary += "  ⚠️ 高波動期：擴大停損間距，降低部位規模";
            else if (pct < 25)
                summary += "  💡 低波動期：可縮小停損，提高資金效率";

            return new VolatilityProfile
            {
                CurrentATR = rawAtr,
                AdaptiveATR = adaptAtr,
                VolPercentile = pct,
                VolRegime = regime,
                AtrMultiplier = mult,
                SuggestedStop = entryPrice > 0 ? adaptAtr * mult : 0,
                Summary = summary
            };
        }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  ⑥ PerformanceAttributionEngine
    //  勝率歸因 + 進場品質評分 + 持倉時間分析
    // ════════════════════════════════════════════════════════════════════════════
    public static class PerformanceAttributionEngine
    {
        // ── 績效分解（三效應分解）──────────────────────────────────────────────
        /// <summary>
        /// 把總報酬分解為：標的選擇 + 進場時機 + 倉位管理 三部分貢獻。
        /// 方法：比較實際結果 vs 等權持倉基準 vs 隨機進場基準。
        /// </summary>
        // Analyze 是 DecomposePerformance 的別名（向前相容）
        public static PerformanceAttribution Analyze(List<TradeJournalEntry> trades)
            => DecomposePerformance(trades);

        public static PerformanceAttribution DecomposePerformance(
            List<TradeJournalEntry> trades)
        {
            if (trades == null || trades.Count == 0)
                return new PerformanceAttribution { Summary = "無交易記錄" };

            var closed = trades.Where(t => t.ExitPrice > 0).ToList();
            if (closed.Count < 3)
                return new PerformanceAttribution { Summary = $"已結算交易 {closed.Count} 筆（需至少 3 筆才能歸因）" };

            // 1. 標的選擇效應：各標的平均報酬率 vs 整體均值的差異
            var byTicker = closed.GroupBy(t => t.Ticker)
                .ToDictionary(g => g.Key,
                    g => g.Average(t => t.ReturnPct));
            double overallMean = closed.Average(t => t.ReturnPct);
            double selectionEffect = byTicker.Values.Average() - overallMean;

            // 2. 進場時機效應：上漲趨勢 vs 下跌趨勢時進場的勝率差
            var wins = closed.Where(t => t.ReturnPct > 0).ToList();
            var losses = closed.Where(t => t.ReturnPct <= 0).ToList();
            double timingEffect = wins.Count > 0 && losses.Count > 0
                ? wins.Average(t => t.ReturnPct) * wins.Count / closed.Count -
                  Math.Abs(losses.Average(t => t.ReturnPct)) * losses.Count / closed.Count
                : 0;

            // 3. 倉位管理效應：報酬率 × 數量 vs 等量持倉
            double avgQty = closed.Average(t => t.Quantity);
            double actualPnL = closed.Sum(t => t.PnL);
            double equalPnL = closed.Sum(t => t.ReturnPct * t.EntryPrice * avgQty);
            double sizingEffect = equalPnL != 0 ? (actualPnL - equalPnL) / Math.Abs(equalPnL) : 0;

            // 4. 風報比 ≥ 2 的勝率
            var highRR = closed.Where(t => t.RiskRewardRatio >= 2).ToList();
            double winRateByRR = highRR.Count > 0
                ? (double)highRR.Count(t => t.ReturnPct > 0) / highRR.Count
                : 0;

            // 5. 持倉時間分析
            var holdingBias = AnalyzeHoldingPeriodBias(closed);

            // 6. 平均進場品質分
            double avgEQ = closed.Average(t => CalcEntryQualityScore(t).Score);

            string summary = $"歸因分析（{closed.Count} 筆）：" +
                             $"選股 {selectionEffect:+0.0%;-0.0%}  " +
                             $"時機 {timingEffect:+0.0%;-0.0%}  " +
                             $"倉管 {sizingEffect:+0.0%;-0.0%}  " +
                             $"進場品質 {avgEQ:F0}/100";

            return new PerformanceAttribution
            {
                SelectionEffect = selectionEffect,
                TimingEffect = timingEffect,
                SizingEffect = sizingEffect,
                TotalEffect = selectionEffect + timingEffect + sizingEffect,
                AvgEntryQuality = avgEQ,
                WinRateByRR = winRateByRR,
                HoldingPeriodBias = holdingBias,
                Summary = summary
            };
        }

        // ── 進場品質評分 ─────────────────────────────────────────────────────
        /// <summary>
        /// 根據 AI 建議 + 進場時指標打分（0~100）。
        /// 以交易日誌的 AiSuggestion 和 ReturnPct 推算：
        ///   RSI 30-50 進場買入 → 高分；RSI>70 追高 → 低分
        /// </summary>
        public static EntryQualityScore CalcEntryQualityScore(TradeJournalEntry t)
        {
            double score = 50;  // 基礎分 50
            string rsiSignal = "-", macdSignal = "-", patternSignal = "-";

            // 從 Notes 欄位取 RSI（如果有）
            double rsi = 0;
            var rsiMatch = System.Text.RegularExpressions.Regex.Match(
                t.Notes ?? "", @"RSI[=:\s]+([\d.]+)");
            if (rsiMatch.Success) double.TryParse(rsiMatch.Groups[1].Value,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out rsi);

            if (rsi > 0)
            {
                if (t.Direction == "Buy")
                {
                    if (rsi < 30) { score += 20; rsiSignal = "超賣買入 +20"; }
                    else if (rsi < 50) { score += 10; rsiSignal = "低位買入 +10"; }
                    else if (rsi > 70) { score -= 15; rsiSignal = "超買追入 -15"; }
                }
                else
                {
                    if (rsi > 70) { score += 20; rsiSignal = "超買放空 +20"; }
                    else if (rsi > 50) { score += 10; rsiSignal = "高位放空 +10"; }
                    else if (rsi < 30) { score -= 15; rsiSignal = "超賣追空 -15"; }
                }
            }

            // 實際結果反推品質
            if (t.ReturnPct > 0.05) { score += 15; patternSignal = $"獲利 {t.ReturnPct:P1} +15"; }
            else if (t.ReturnPct > 0.02) { score += 8; patternSignal = $"小獲利 {t.ReturnPct:P1} +8"; }
            else if (t.ReturnPct < -0.05) { score -= 20; patternSignal = $"大虧損 {t.ReturnPct:P1} -20"; }
            else if (t.ReturnPct < -0.02) { score -= 10; patternSignal = $"小虧損 {t.ReturnPct:P1} -10"; }

            // 風報比加分
            if (t.RiskRewardRatio >= 2) { score += 10; macdSignal = $"RR={t.RiskRewardRatio:F1} +10"; }

            score = Math.Max(0, Math.Min(100, score));
            string grade = score >= 80 ? "A" : score >= 65 ? "B" : score >= 50 ? "C" : "D";

            return new EntryQualityScore
            {
                TradeId = t.Id,
                Score = score,
                Grade = grade,
                RsiSignal = rsiSignal,
                MacdSignal = macdSignal,
                PatternSignal = patternSignal,
                Remark = $"{grade} 級  {score:F0}分"
            };
        }

        // ── 持倉時間 → 勝率相關性 ─────────────────────────────────────────────
        public static Dictionary<string, double> AnalyzeHoldingPeriodBias(
            List<TradeJournalEntry> trades)
        {
            var result = new Dictionary<string, double>();
            var closed = trades.Where(t => t.ExitPrice > 0 && t.TradeDate != default).ToList();
            if (closed.Count < 3) return result;

            // 以 TradeDate 當進場日（簡化：Exit 不在 model 裡，用 Notes 推算或靠 PnL）
            // Bucket by estimated holding period (簡化版用 Id 順序間距估算)
            var buckets = new[]
            {
                ("1-3天",   1, 3),
                ("4-7天",   4, 7),
                ("8-14天",  8, 14),
                ("15-30天", 15, 30),
                (">30天",   31, 9999)
            };

            // 用 TradeDate 和下一筆同 Ticker 的 exit 計算天數（若無 ExitDate 欄位，以 Id 差估計）
            foreach (var (label, dMin, dMax) in buckets)
            {
                var inBucket = closed.Where(t =>
                {
                    // 簡易估算：以 Id 差或同 Ticker 相鄰筆次 TradeDate 差
                    var next = closed.Where(x => x.Ticker == t.Ticker && x.Id > t.Id).FirstOrDefault();
                    int days = next != null ? (int)(next.TradeDate - t.TradeDate).TotalDays : 5;
                    return days >= dMin && days <= dMax;
                }).ToList();

                if (inBucket.Count >= 2)
                    result[label] = inBucket.Average(t => t.ReturnPct);
            }
            return result;
        }
    }

}