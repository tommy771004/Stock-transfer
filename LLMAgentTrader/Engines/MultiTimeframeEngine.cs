using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace LLMAgentTrader
{
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
                    signal.Weekly_KD_K = wLast.KD_K;
                    signal.Weekly_ADX = wLast.DMI_ADX;
                    signal.Weekly_RSI_Div = wLast.RSI_Divergence;
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
                    signal.Daily_KD_K = dLast.KD_K;
                    signal.Daily_ADX = dLast.DMI_ADX;
                    signal.Daily_RSI_Div = dLast.RSI_Divergence;
                    signal.Daily_MACD_Div = dLast.MACD_Divergence;
                    signal.Daily_Pattern2 = dLast.Pattern2;
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
                    signal.Hourly_KD_K = hLast.KD_K;
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

        /// <summary>將多時間框架信號轉換為 AI 提示詞補充（含 KD、ADX、背離、複合型態）</summary>
        public static string ToPromptContext(MultiTimeframeSignal sig)
        {
            if (sig == null) return "";
            var sb = new StringBuilder("\n【多時間框架共振分析（含 KD / ADX / 背離）】\n");

            // 週線
            sb.Append($"週線: {sig.Weekly_Trend} | RSI={sig.Weekly_RSI:F1} | KD_K={sig.Weekly_KD_K:F1}" +
                      $" | ADX={sig.Weekly_ADX:F1} | MACD柱={sig.Weekly_MACD_Hist:F3} | 形態={sig.Weekly_Pattern}");
            if (sig.Weekly_RSI_Div != "-") sb.Append($" | ⚡RSI{sig.Weekly_RSI_Div}");
            sb.AppendLine();

            // 日線
            sb.Append($"日線: {sig.Daily_Trend} | RSI={sig.Daily_RSI:F1} | KD_K={sig.Daily_KD_K:F1}" +
                      $" | ADX={sig.Daily_ADX:F1} | MACD柱={sig.Daily_MACD_Hist:F3} | 單K={sig.Daily_Pattern}");
            if (sig.Daily_Pattern2 != "-") sb.Append($" | 複合={sig.Daily_Pattern2}");
            if (sig.Daily_RSI_Div != "-") sb.Append($" | ⚡RSI{sig.Daily_RSI_Div}");
            if (sig.Daily_MACD_Div != "-") sb.Append($" | ⚡MACD{sig.Daily_MACD_Div}");
            sb.AppendLine();

            // 小時線
            sb.AppendLine($"小時線: {sig.Hourly_Trend} | RSI={sig.Hourly_RSI:F1} | KD_K={sig.Hourly_KD_K:F1}" +
                          $" | MACD柱={sig.Hourly_MACD_Hist:F3} | 形態={sig.Hourly_Pattern}");

            sb.AppendLine(sig.AlignmentSummary);
            sb.AppendLine("👉 請在分析中特別考量多時間框架共振程度，並根據 KD 超賣/超買區、DMI 趨勢強度、" +
                          "RSI/MACD 背離信號及 K 線複合型態（晨星/夜星/三兵等）給出更精準的操作建議。");
            return sb.ToString();
        }
    }
}
