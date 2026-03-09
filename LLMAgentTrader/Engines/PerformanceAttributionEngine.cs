using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LLMAgentTrader
{
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
