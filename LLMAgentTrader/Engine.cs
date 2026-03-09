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
        private static readonly string EncPath        = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "APIKey.enc");
        private static readonly string LegacyPath     = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "APIKey.txt");
        private static readonly string GeminiEncPath  = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GeminiKey.enc");
        private static readonly string GeminiLegacyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GeminiKey.txt");

        // ── OpenRouter Key ────────────────────────────────────────────────────
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
                byte[] enc  = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(EncPath, enc);
            }
            catch (Exception ex) { AppLogger.Log("ApiKeyManager.Save 失敗", ex); }
        }

        // ── Google AI Studio (Gemini) Key ─────────────────────────────────────
        public static string LoadGemini()
        {
            if (File.Exists(GeminiEncPath))
            {
                try
                {
                    byte[] enc = File.ReadAllBytes(GeminiEncPath);
                    byte[] dec = ProtectedData.Unprotect(enc, null, DataProtectionScope.CurrentUser);
                    return Encoding.UTF8.GetString(dec);
                }
                catch (Exception ex) { AppLogger.Log("ApiKeyManager.LoadGemini 解密失敗", ex); }
            }
            // 舊版明文檔案自動遷移
            if (File.Exists(GeminiLegacyPath))
            {
                try
                {
                    string key = File.ReadAllText(GeminiLegacyPath).Trim();
                    if (!string.IsNullOrEmpty(key)) { SaveGemini(key); File.Delete(GeminiLegacyPath); return key; }
                }
                catch (Exception ex) { AppLogger.Log("ApiKeyManager.MigrateGemini 失敗", ex); }
            }
            return "";
        }

        public static void SaveGemini(string key)
        {
            try
            {
                byte[] data = Encoding.UTF8.GetBytes(key);
                byte[] enc  = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(GeminiEncPath, enc);
            }
            catch (Exception ex) { AppLogger.Log("ApiKeyManager.SaveGemini 失敗", ex); }
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

            // VWAP（每個交易日重置，VWAP 必須以單日為計算單位）
            double cumVol = 0, cumPv = 0;
            for (int i = 0; i < list.Count; i++)
            {
                // 偵測跨日：日期改變時重置累積量
                if (i > 0 && list[i].Date.Date != list[i - 1].Date.Date)
                {
                    cumVol = 0;
                    cumPv = 0;
                }
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

            // ATR 正確初始化：使用前14根 True Range 的簡單平均作為 Wilder 種子值
            // 若從 0 開始，前20根 ATR 嚴重偏低，導致停損設置過窄
            double atr = 0;
            if (list.Count >= 15)
            {
                double sumTr14 = 0;
                for (int k = 1; k <= 14; k++)
                    sumTr14 += Math.Max(list[k].High - list[k].Low,
                               Math.Max(Math.Abs(list[k].High - list[k - 1].Close),
                                        Math.Abs(list[k].Low - list[k - 1].Close)));
                atr = sumTr14 / 14.0;
                list[14].ATR = atr;
            }

            // ATR (Wilder SMMA) + 布林通道 (母體標準差 ÷n) + K線型態
            for (int i = 1; i < list.Count; i++)
            {
                double tr = Math.Max(list[i].High - list[i].Low,
                            Math.Max(Math.Abs(list[i].High - list[i - 1].Close),
                                     Math.Abs(list[i].Low - list[i - 1].Close)));
                // i=14 已在初始化時設定，i≥15 才使用 Wilder 平滑
                if (i >= 15)
                {
                    atr = (atr * 13 + tr) / 14;
                    list[i].ATR = atr;
                }
                else if (i < 14)
                {
                    // 前13根不計算 ATR，等待第14根初始化完成
                    list[i].ATR = 0;
                }

                if (i >= 20)
                {
                    // 布林通道使用母體標準差（÷n=20），與 TradingView/MT5 一致
                    var sl = list.Skip(i - 19).Take(20).Select(x => x.Close).ToList();
                    double avg = sl.Average();
                    double sd = Math.Sqrt(sl.Select(x => Math.Pow(x - avg, 2)).Sum() / 20.0);
                    list[i].BB_Middle = avg; list[i].BB_Upper = avg + sd * 2; list[i].BB_Lower = avg - sd * 2;
                    list[i].BB_Width = avg > 0 ? (list[i].BB_Upper - list[i].BB_Lower) / avg * 100 : 0;
                }

                double body = Math.Abs(list[i].Open - list[i].Close), range = list[i].High - list[i].Low;
                double upperShadow = list[i].High - Math.Max(list[i].Open, list[i].Close);
                double lowerShadow = Math.Min(list[i].Open, list[i].Close) - list[i].Low;
                bool isBull = list[i].Close >= list[i].Open;
                bool isPrevBull = list[i - 1].Close >= list[i - 1].Open;
                double prevBody = Math.Abs(list[i - 1].Open - list[i - 1].Close);

                // ── 單根K棒型態 ────────────────────────────────────────────────
                if (range > 0 && body / range < 0.1 && upperShadow > range * 0.45 && lowerShadow < range * 0.1)
                    list[i].Pattern = "墓碑線";   // Gravestone Doji (空頭反轉)
                else if (range > 0 && body / range < 0.1)
                    list[i].Pattern = "十字星";   // Doji
                else if (body > 0 && lowerShadow >= body * 2 && upperShadow <= body * 0.3 && !isBull)
                    list[i].Pattern = "槌子";        // Hammer (看漲，陰線槌更強)
                else if (body > 0 && upperShadow >= body * 2 && lowerShadow <= body * 0.3 && !isBull)
                    list[i].Pattern = "倒槌子";      // Inverted Hammer (潛在看漲)
                else if (body > 0 && lowerShadow >= body * 2 && upperShadow <= body * 0.3 && isBull)
                    list[i].Pattern = "吊人線";      // Hanging Man (看跌警示)
                else if (body > 0 && upperShadow >= body * 2 && lowerShadow <= body * 0.3 && isBull)
                    list[i].Pattern = "流星";        // Shooting Star (看跌)
                else if (range > 0 && upperShadow > range * 0.5 && body < range * 0.25)
                    list[i].Pattern = "上影線";   // Long Upper Shadow
                else if (range > 0 && lowerShadow > range * 0.5 && body < range * 0.25)
                    list[i].Pattern = "下影線";   // Long Lower Shadow
                // ── 兩根K棒型態 ──────────────────────────────────────────────
                else if (isBull && !isPrevBull && list[i].Close > list[i - 1].Open && list[i].Open < list[i - 1].Close)
                    list[i].Pattern = "多頭吞噬"; // Bullish Engulfing
                else if (!isBull && isPrevBull && list[i].Close < list[i - 1].Open && list[i].Open > list[i - 1].Close)
                    list[i].Pattern = "空頭吞噬"; // Bearish Engulfing
                else if (!isPrevBull && isBull && prevBody > 0 && body < prevBody * 0.6 &&
                         list[i].Close <= list[i - 1].Open && list[i].Open >= list[i - 1].Close)
                    list[i].Pattern = "多頭母子"; // Bullish Harami
                else if (isPrevBull && !isBull && prevBody > 0 && body < prevBody * 0.6 &&
                         list[i].Close >= list[i - 1].Open && list[i].Open <= list[i - 1].Close)
                    list[i].Pattern = "空頭母子"; // Bearish Harami
                else
                    list[i].Pattern = "-";

                // ── 三根K棒型態 (存入 Pattern2) ───────────────────────────────
                if (i >= 2)
                {
                    var p2 = list[i - 2]; var p1 = list[i - 1]; var p0 = list[i];
                    double b2 = Math.Abs(p2.Close - p2.Open);
                    double b1 = Math.Abs(p1.Close - p1.Open);
                    double b0 = Math.Abs(p0.Close - p0.Open);
                    bool bull2 = p2.Close >= p2.Open, bull1 = p1.Close >= p1.Open, bull0 = p0.Close >= p0.Open;

                    // 晨星 (Morning Star): 大黑K + 小實體 + 大紅K (多頭反轉)
                    if (!bull2 && b2 > 0 && b1 < b2 * 0.5 && bull0 && b0 > b2 * 0.5 &&
                        p0.Close > (p2.Open + p2.Close) / 2)
                        list[i].Pattern2 = "晨星";
                    // 夜星 (Evening Star): 大紅K + 小實體 + 大黑K (空頭反轉)
                    else if (bull2 && b2 > 0 && b1 < b2 * 0.5 && !bull0 && b0 > b2 * 0.5 &&
                             p0.Close < (p2.Open + p2.Close) / 2)
                        list[i].Pattern2 = "夜星";
                    // 紅三兵 (Three White Soldiers): 連續三根多頭K棒，逐漸墊高
                    else if (bull0 && bull1 && bull2 &&
                             p0.Close > p1.Close && p1.Close > p2.Close &&
                             p0.Open > p1.Open && p1.Open > p2.Open &&
                             b0 > 0 && b1 > 0 && b2 > 0)
                        list[i].Pattern2 = "紅三兵";
                    // 黑三兵 (Three Black Crows): 連續三根空頭K棒，逐漸墜低
                    else if (!bull0 && !bull1 && !bull2 &&
                             p0.Close < p1.Close && p1.Close < p2.Close &&
                             p0.Open < p1.Open && p1.Open < p2.Open &&
                             b0 > 0 && b1 > 0 && b2 > 0)
                        list[i].Pattern2 = "黑三兵";
                    else
                        list[i].Pattern2 = "-";
                }
            }

            // ── SMA (5/10/20/60) + 乖離率 + 布林緊縮  改用滾動累加 O(n) ────────
            {
                double sum5 = 0, sum10 = 0, sum20 = 0, sum60 = 0;
                for (int i = 0; i < list.Count; i++)
                {
                    double c = list[i].Close;
                    sum5  += c; sum10 += c; sum20 += c; sum60 += c;

                    if (i >= 5)  sum5  -= list[i - 5].Close;
                    if (i >= 10) sum10 -= list[i - 10].Close;
                    if (i >= 20) sum20 -= list[i - 20].Close;
                    if (i >= 60) sum60 -= list[i - 60].Close;

                    if (i >= 4)  list[i].SMA5  = sum5  / 5;
                    if (i >= 9)  list[i].SMA10 = sum10 / 10;
                    if (i >= 19) list[i].SMA20 = sum20 / 20;
                    if (i >= 59) list[i].SMA60 = sum60 / 60;

                    if (list[i].SMA20 > 0)
                        list[i].Bias20 = (list[i].Close - list[i].SMA20) / list[i].SMA20 * 100;
                    list[i].BB_Squeeze = list[i].BB_Width > 0 && list[i].BB_Width < 3.5;
                }
            }

            // ── KD 隨機指標 (Fast K=9, Slow D=3) ──────────────────────────────
            // %K = (Close - LowestLow9) / (HighestHigh9 - LowestLow9) × 100
            // %D = 3日 SMA of %K
            {
                var kVals = new double[list.Count];
                for (int i = 8; i < list.Count; i++)
                {
                    var w = list.Skip(i - 8).Take(9).ToList();
                    double lo = w.Min(x => x.Low), hi = w.Max(x => x.High);
                    kVals[i] = (hi - lo) > 0 ? (list[i].Close - lo) / (hi - lo) * 100 : 50;
                    list[i].KD_K = kVals[i];
                    if (i >= 10) list[i].KD_D = (kVals[i] + kVals[i - 1] + kVals[i - 2]) / 3.0;
                }
            }

            // ── DMI / ADX (14期 Wilder) ─────────────────────────────────────────
            // +DM, -DM, TR → Wilder平滑14 → +DI, -DI → DX → ADX
            if (list.Count >= 28)
            {
                double smPlus = 0, smMinus = 0, smTR = 0, adxSmooth = 0;
                bool adxSeeded = false;
                double adxSum = 0; int adxCount = 0;

                // 初始化前14根
                for (int i = 1; i <= 14; i++)
                {
                    double hi = list[i].High, lo = list[i].Low;
                    double pHi = list[i - 1].High, pLo = list[i - 1].Low, pCl = list[i - 1].Close;
                    double dmPlus = (hi - pHi) > (pLo - lo) && (hi - pHi) > 0 ? hi - pHi : 0;
                    double dmMinus = (pLo - lo) > (hi - pHi) && (pLo - lo) > 0 ? pLo - lo : 0;
                    double tr = Math.Max(hi - lo, Math.Max(Math.Abs(hi - pCl), Math.Abs(lo - pCl)));
                    smPlus += dmPlus; smMinus += dmMinus; smTR += tr;
                }

                for (int i = 15; i < list.Count; i++)
                {
                    double hi = list[i].High, lo = list[i].Low;
                    double pHi = list[i - 1].High, pLo = list[i - 1].Low, pCl = list[i - 1].Close;
                    double dmPlus = (hi - pHi) > (pLo - lo) && (hi - pHi) > 0 ? hi - pHi : 0;
                    double dmMinus = (pLo - lo) > (hi - pHi) && (pLo - lo) > 0 ? pLo - lo : 0;
                    double tr = Math.Max(hi - lo, Math.Max(Math.Abs(hi - pCl), Math.Abs(lo - pCl)));

                    smPlus  = smPlus  - smPlus  / 14 + dmPlus;
                    smMinus = smMinus - smMinus / 14 + dmMinus;
                    smTR    = smTR    - smTR    / 14 + tr;

                    double diPlus  = smTR > 0 ? smPlus  / smTR * 100 : 0;
                    double diMinus = smTR > 0 ? smMinus / smTR * 100 : 0;
                    list[i].DMI_Plus = diPlus;
                    list[i].DMI_Minus = diMinus;

                    double dx = (diPlus + diMinus) > 0
                        ? Math.Abs(diPlus - diMinus) / (diPlus + diMinus) * 100 : 0;

                    if (!adxSeeded)
                    {
                        adxSum += dx; adxCount++;
                        if (adxCount == 14) { adxSmooth = adxSum / 14; adxSeeded = true; list[i].DMI_ADX = adxSmooth; }
                    }
                    else
                    {
                        adxSmooth = (adxSmooth * 13 + dx) / 14;
                        list[i].DMI_ADX = adxSmooth;
                    }
                }
            }

            // ── RSI 背離 + MACD 背離 偵測 ──────────────────────────────────────
            DivergenceDetector.Detect(list);
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
            public double VaR95 { get; set; }   // 95% 單日最大虧損（正值表示虧損）
            public double CVaR95 { get; set; }  // 95% 條件風險值（Tail Risk）
            public double AvgInvestFraction { get; set; }  // 平均投入比例（Kelly動態調整後）
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

            // 修正 Look-ahead Bias：使用「前一根」K棒的 AgentAction 在「當根」執行
            // 這樣 AI 是基於前一根收盤的指標決策，在下一根開盤/收盤執行，符合實際交易邏輯
            for (int idx = 1; idx < data.Count; idx++)
            {
                var signal = data[idx - 1]; // 前一根 K 棒的信號
                var d = data[idx];          // 當根 K 棒的執行價格

                if (signal.AgentAction == "Buy" && pos == 0 && d.Close > 0)
                {
                    // ── 動態 Kelly 倉位（基於累積交易績效）────────────────────────────
                    double dynamicKelly = 0.25; // 預設 quarter-kelly
                    if (trades >= 5 && winPcts.Count + lossPcts.Count >= 5)
                    {
                        double w = winPcts.Count > 0 ? winPcts.Average() : 0;
                        double l = lossPcts.Count > 0 ? Math.Abs(lossPcts.Average()) : 0.01;
                        double wr = (double)wins / trades;
                        double kFull = (w > 0 && l > 0) ? wr - (1 - wr) / (w / l) : 0;
                        kFull = Math.Max(0, kFull);
                        // Half-Kelly 縮減（更保守）
                        dynamicKelly = kFull * 0.5;
                        dynamicKelly = Math.Max(0.05, Math.Min(0.40, dynamicKelly));
                    }
                    double invest = cap * dynamicKelly;
                    double costPerShare = d.Close * (1 + BuyFee);
                    double lots = Math.Floor(invest / (costPerShare * 1000));
                    if (lots >= 1) { pos = lots * 1000; cap -= pos * costPerShare; entry = d.Close; trades++; }
                }
                else if (signal.AgentAction == "Sell" && pos > 0)
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

            // Sortino Ratio（下行標準差門檻使用 0，不是無風險利率）
            // 標準 Sortino 定義：只計算報酬率 < 0 的日子，以 0 為損失門檻
            // 使用 rfDaily 作門檻會高估下行風險，使 Sortino 比率偏低
            double sortino = 0;
            if (dailyReturns.Count > 1)
            {
                double meanR = dailyReturns.Average();
                var downside = dailyReturns.Where(r => r < 0).ToList();
                if (downside.Count > 1)
                {
                    // 下行半變異數（Downside Semi-Deviation）：以 0 為門檻
                    double downStd = Math.Sqrt(downside.Select(r => r * r).Average());
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
                // 修正：同時檢查 w > 0 且 l > 0，避免 w=0 時 (w/l)=0 造成除以零 → NaN
                kelly = (w > 0 && l > 0) ? winRate - (1 - winRate) / (w / l) : 0;
                kelly = Math.Max(0, Math.Min(0.5, kelly)); // 限制在 0~50%
            }

            // VaR 95%：在信心水準 95% 下，單日最壞的損失是多少
            // CVaR 95%（Expected Shortfall）：尾部最差 5% 日子的平均損失
            double var95 = 0, cvar95 = 0;
            if (dailyReturns.Count > 5)
            {
                var sorted = dailyReturns.OrderBy(r => r).ToList();
                int cutoff = (int)Math.Floor(sorted.Count * 0.05);
                if (cutoff < 1) cutoff = 1;
                var95 = -sorted[cutoff - 1];   // 取第 5 百分位（負報酬轉為正的損失數字）
                cvar95 = -sorted.Take(cutoff).Average();
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
                AvgLossPct = lossPcts.Count > 0 ? Math.Abs(lossPcts.Average()) : 0,
                VaR95 = var95,
                CVaR95 = cvar95,
                AvgInvestFraction = 0
            };
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
            CancellationToken ct = default,
            InstitutionalData institutional = null,
            MarginData margin = null)
        {
            var target = data.TakeLast(30).ToList();
            var sb = new StringBuilder($"【近期市場新聞情緒與基本面估值】\n{newsContext}");

            if (mtfSignal != null)
                sb.Append(MultiTimeframeEngine.ToPromptContext(mtfSignal));

            // ── 台股籌碼面（三大法人 + 融資融券）────────────────────────────
            string instCtx   = InstitutionalService.ToPromptContext(institutional);
            string marginCtx = MarginTradingService.ToPromptContext(margin);
            if (!string.IsNullOrEmpty(instCtx))   sb.Append("\n" + instCtx);
            if (!string.IsNullOrEmpty(marginCtx)) sb.Append(marginCtx);

            // 輸出最新一根K棒的進階指標摘要
            var latest = target.Last();
            sb.Append("\n\n【進階技術指標摘要（最新K棒）】\n");
            sb.AppendLine($"KD: K={latest.KD_K:F1} / D={latest.KD_D:F1}" +
                          (latest.KD_K < 20 ? " ⚡超賣區" : latest.KD_K > 80 ? " 🔥超買區" : ""));
            sb.AppendLine($"DMI: +DI={latest.DMI_Plus:F1} / -DI={latest.DMI_Minus:F1} / ADX={latest.DMI_ADX:F1}" +
                          (latest.DMI_ADX > 25 ? (latest.DMI_Plus > latest.DMI_Minus ? " 強多頭趨勢" : " 強空頭趨勢") : " 盤整無趨勢"));
            sb.AppendLine($"SMA均線: 5={latest.SMA5:F1} / 10={latest.SMA10:F1} / 20={latest.SMA20:F1} / 60={latest.SMA60:F1}");
            sb.AppendLine($"乖離率(SMA20): {latest.Bias20:F2}%" + (latest.Bias20 > 5 ? " (偏高，注意回調)" : latest.Bias20 < -5 ? " (偏低，注意反彈)" : ""));
            if (latest.BB_Squeeze) sb.AppendLine("⚡ 布林通道緊縮（BB Width < 3.5%），醞釀方向性突破");
            if (latest.RSI_Divergence != "-") sb.AppendLine($"⚡ RSI {latest.RSI_Divergence}（背離信號）");
            if (latest.MACD_Divergence != "-") sb.AppendLine($"⚡ MACD {latest.MACD_Divergence}（背離信號）");
            if (latest.Pattern2 != "-") sb.AppendLine($"⚡ 複合K線型態: {latest.Pattern2}");

            sb.Append("\n【歷史技術數據】\nDate,Close,RSI,KD_K,KD_D,MACD_Hist,BBW,+DI,-DI,ADX,Pattern,Pattern2,RSI_Div,MACD_Div\n");
            foreach (var d in target)
                sb.AppendLine($"{d.Date:MMdd},{d.Close:F1},{d.RSI:F1},{d.KD_K:F1},{d.KD_D:F1},{d.MACD_Hist:F3},{d.BB_Width:F1},{d.DMI_Plus:F1},{d.DMI_Minus:F1},{d.DMI_ADX:F1},{d.Pattern},{d.Pattern2},{d.RSI_Divergence},{d.MACD_Divergence}");

            ct.ThrowIfCancellationRequested();

            // ── JSON schema 要求（附在 system prompt 末尾，兩條路徑都用）────────
            string sysPrompt = prompt +
                "\n🚨請務必用「繁體中文」回覆。\n" +
                "輸出 JSON: { \"debate_log\": \"...\", \"results\": [ { \"Date\": \"MMdd\", \"Action\": \"Buy/Sell/Hold\", \"Reasoning\": \"...\" } ] }";
            string userContent = sb.ToString();

            string raw;

            if (LlmConfig.IsGeminiDirect(LlmConfig.CurrentModel))
            {
                // ── Google AI Studio 直連路徑 ─────────────────────────────────
                string geminiKey = !string.IsNullOrEmpty(key) ? key : LlmConfig.GeminiApiKey;
                raw = await GeminiService.CallAsync(
                    geminiKey,
                    LlmConfig.GeminiModelName(LlmConfig.CurrentModel),
                    sysPrompt, userContent,
                    jsonMode: true, ct);
            }
            else
            {
                // ── OpenRouter 路徑（原有邏輯）───────────────────────────────
                var payload = new
                {
                    model = LlmConfig.CurrentModel,
                    messages = new[]
                    {
                        new { role = "system", content = sysPrompt },
                        new { role = "user", content = userContent }
                    },
                    response_format = new { type = "json_object" }
                };
                var req = new HttpRequestMessage(HttpMethod.Post, LlmConfig.BaseUrl)
                {
                    Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
                };
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", key);
                var resp = await AppHttpClients.Llm.SendAsync(req, ct);
                resp.EnsureSuccessStatusCode();
                string outerRaw = await resp.Content.ReadAsStringAsync();
                raw = JsonDocument.Parse(outerRaw).RootElement
                          .GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString();
            }

            var root = JsonDocument.Parse(raw).RootElement;

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

}