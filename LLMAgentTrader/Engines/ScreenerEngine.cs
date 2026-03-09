using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LLMAgentTrader
{
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

                    // ── 新增：KD 篩選 ────────────────────────────────────────────
                    bool kdCrossUp = data.Count >= 3 &&
                        data[data.Count - 2].KD_K < data[data.Count - 2].KD_D &&
                        last.KD_K >= last.KD_D;
                    if (criteria.KD_Oversold && last.KD_K < 20) matched.Add($"KD超賣 K={last.KD_K:F1}");
                    else if (criteria.KD_Oversold) continue;
                    if (criteria.KD_Overbought && last.KD_K > 80) matched.Add($"KD超買 K={last.KD_K:F1}");
                    else if (criteria.KD_Overbought) continue;
                    if (criteria.KD_CrossUp && kdCrossUp) matched.Add("KD黃金交叉");
                    else if (criteria.KD_CrossUp) continue;

                    // ── 新增：DMI 篩選 ────────────────────────────────────────────
                    if (criteria.DMI_Bullish && last.DMI_Plus > last.DMI_Minus && last.DMI_ADX > 25)
                        matched.Add($"DMI強多頭 +DI={last.DMI_Plus:F1} ADX={last.DMI_ADX:F1}");
                    else if (criteria.DMI_Bullish) continue;

                    // ── 新增：布林緊縮 ────────────────────────────────────────────
                    if (criteria.BB_Squeeze && last.BB_Squeeze) matched.Add($"布林緊縮 BBW={last.BB_Width:F1}%");
                    else if (criteria.BB_Squeeze) continue;

                    // ── 新增：背離信號 ────────────────────────────────────────────
                    if (criteria.RSI_BullDiv && last.RSI_Divergence == "正背離") matched.Add("RSI正背離");
                    else if (criteria.RSI_BullDiv) continue;
                    if (criteria.MACD_BullDiv && last.MACD_Divergence == "底背離") matched.Add("MACD底背離");
                    else if (criteria.MACD_BullDiv) continue;

                    // ── 新增：SMA 多頭排列 ────────────────────────────────────────
                    bool smaArrange = last.SMA5 > 0 && last.SMA10 > 0 && last.SMA20 > 0 && last.SMA60 > 0 &&
                                     last.SMA5 > last.SMA10 && last.SMA10 > last.SMA20 && last.SMA20 > last.SMA60;
                    if (criteria.SMA_BullArrange && smaArrange) matched.Add("SMA多頭排列");
                    else if (criteria.SMA_BullArrange) continue;

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
                        MatchedRules = matched,
                        KD_K = last.KD_K,
                        KD_D = last.KD_D,
                        DMI_ADX = last.DMI_ADX,
                        DMI_Plus = last.DMI_Plus,
                        DMI_Minus = last.DMI_Minus,
                        Bias20 = last.Bias20,
                        RSI_Divergence = last.RSI_Divergence,
                        MACD_Divergence = last.MACD_Divergence,
                        Pattern2 = last.Pattern2
                    });
                }
                catch (Exception ex) { AppLogger.Log($"ScreenerEngine.ScanAsync {ticker} 失敗", ex); }
            }

            return results.OrderByDescending(r => r.MatchScore).ToList();
        }
    }
}
