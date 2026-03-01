using System;
using System.Collections.Generic;
using System.Linq;

namespace LLMAgentTrader
{
    /// <summary>
    /// 位置管理 / 倉位大小計算服務
    /// 提供 Kelly、固定百分比、金字塔分批等常用方法。
    /// </summary>
    public static class PositionSizingService
    {
        /// <summary>
        /// 計算 Kelly fraction: f* = W - (1-W)/R
        /// winRate: 勝率 (0..1)
        /// avgWin: 平均勝率（以報酬率表示，例如 0.05 表示 5%）
        /// avgLoss: 平均虧損（正值，例如 0.03 表示 -3%）
        /// 回傳值會限制到 [0, 0.5]，避免過度建議
        /// </summary>
        public static double CalcKellyFraction(double winRate, double avgWin, double avgLoss)
        {
            if (winRate <= 0 || avgWin <= 0 || avgLoss <= 0) return 0.0;
            double R = avgWin / avgLoss;
            if (R <= 0) return 0.0;
            double k = winRate - (1.0 - winRate) / R;
            // 限制範圍：0 ~ 0.5
            k = Math.Max(0.0, k);
            k = Math.Min(0.5, k); // 以 50% 做為上限（實務常做縮減）
            return k;
        }

        /// <summary>
        /// 從交易日誌計算 Kelly（使用已平倉交易）
        /// 若無足夠交易回傳 0
        /// </summary>
        public static double CalcKellyFromJournalEntries(IEnumerable<TradeJournalEntry> entries)
        {
            if (entries == null) return 0.0;
            var closed = entries.Where(e => e.ExitPrice > 0).ToList();
            if (closed.Count < 5) return 0.0; // 樣本太少時不建議使用
            int wins = 0;
            var winPct = new List<double>();
            var lossPct = new List<double>();
            foreach (var t in closed)
            {
                if (t.EntryPrice <= 0) continue;
                double ret = (t.ExitPrice - t.EntryPrice) / t.EntryPrice * (t.Direction == "Buy" ? 1.0 : -1.0);
                if (ret > 0) { wins++; winPct.Add(ret); }
                else if (ret < 0) lossPct.Add(Math.Abs(ret));
            }
            if (winPct.Count == 0 || lossPct.Count == 0) return 0.0;
            double w = (double)wins / closed.Count;
            double avgW = winPct.Average();
            double avgL = lossPct.Average();
            return CalcKellyFraction(w, avgW, avgL);
        }

        /// <summary>
        /// 由 Kelly fraction 與 shrink 參數計算實際建議倉位比率
        /// shrink 建議預設 0.5（即 half-kelly）
        /// </summary>
        public static double CalcPositionFractionKelly(double kellyFraction, double shrink = 0.5)
        {
            if (kellyFraction <= 0) return 0.0;
            shrink = Math.Clamp(shrink, 0.0, 1.0);
            return Math.Max(0.0, Math.Min(1.0, kellyFraction * shrink));
        }

        /// <summary>
        /// 根據資金與 fraction 計算建議股數（會向下取整至 lotSize 的倍數）
        /// fraction: 想承擔的風險資本比率（例如 0.02 表示承擔資金的2%風險）
        /// entryPrice / stopPrice: 進場價、停損價
        /// lotSize: 整批最小單位（例如 TW=1000；美股常為 1）
        /// 備註：若停損等於進場或計算出數量小於一個 lot，回傳 0
        /// </summary>
        public static long CalcQuantityFromFraction(double capital, double fraction, double entryPrice, double stopPrice, int lotSize = 1)
        {
            if (capital <= 0 || fraction <= 0 || entryPrice <= 0 || lotSize <= 0) return 0;
            double riskPerShare = Math.Abs(entryPrice - stopPrice);
            if (riskPerShare <= 1e-9) return 0;
            double riskBudget = capital * fraction;
            double rawShares = riskBudget / riskPerShare;
            if (rawShares < 1) return 0;

            // 先向下對齊到 lotSize
            long qty = NormalizeToLot((long)Math.Floor(rawShares), lotSize);
            if (qty < lotSize) return 0;
            return qty;
        }

        /// <summary>
        /// 固定百分比倉位計算（直接以資金百分比換算）
        /// percent: 例如 0.02 表示以 2% 資金投入風險
        /// 若提供 stopPrice 則以風險金額計算股數（同 CalcQuantityFromFraction），否則以投入金額 / entryPrice 計算股數
        /// </summary>
        public static long CalcQuantityFixedPercent(double capital, double percent, double entryPrice, double? stopPrice = null, int lotSize = 1)
        {
            if (lotSize <= 0) return 0;
            if (stopPrice.HasValue)
            {
                return CalcQuantityFromFraction(capital, percent, entryPrice, stopPrice.Value, lotSize);
            }
            if (capital <= 0 || percent <= 0 || entryPrice <= 0) return 0;
            double invest = capital * percent;
            long raw = (long)Math.Floor(invest / entryPrice);
            long qty = NormalizeToLot(raw, lotSize);
            if (qty < lotSize) return 0;
            return qty;
        }

        /// <summary>
        /// 產生金字塔分批進場陣列
        /// totalQty: 總股數（已是整股），若非 lotSize 倍數會向下調整
        /// levels: 分成幾個批次（>=1）
        /// stepPct: 每層價格步幅（例如 0.01 = 1%），對做多而言每層價格 = entryPrice - i*stepPct*entryPrice
        /// up: 若 true 表示做多（價格從低到高），若 false 表示做空（價格從高到低）
        /// 回傳 (price, qty) 的列表，qty 已調整為 lotSize 倍數
        /// </summary>
        public static List<(double Price, long Qty)> GeneratePyramidEntries(double entryPrice, long totalQty, int levels = 3, double stepPct = 0.01, int lotSize = 1, bool up = true)
        {
            var outList = new List<(double, long)>();
            if (levels <= 0 || totalQty <= 0 || entryPrice <= 0 || lotSize <= 0) return outList;

            // 強制 totalQty 為 lotSize 倍數（向下）
            totalQty = NormalizeToLot(totalQty, lotSize);
            if (totalQty <= 0) return outList;

            // 權重採 1,2,...,levels (較晚加碼佔更多)
            var weights = Enumerable.Range(1, levels).Select(i => (double)i).ToArray();
            double sumW = weights.Sum();

            // 初始分配（未對齊 lot）
            var rawQtys = weights.Select(w => totalQty * (w / sumW)).ToArray();

            // 向下取整到 lotSize 並保留殘差用於再次分配
            var finalQtys = new long[levels];
            long assigned = 0;
            for (int i = 0; i < levels; i++)
            {
                long q = NormalizeToLot((long)Math.Floor(rawQtys[i]), lotSize);
                finalQtys[i] = q;
                assigned += q;
            }

            // 分配剩餘的 lot (以從第一層到最後層的順序分配一個 lot，直到用完或無空間)
            long remaining = totalQty - assigned;
            int idx = 0;
            while (remaining >= lotSize)
            {
                finalQtys[idx % levels] += lotSize;
                remaining -= lotSize;
                idx++;
            }

            // 建立價格與數量輸出（維持層級順序）
            for (int i = 0; i < levels; i++)
            {
                double price;
                if (up)
                {
                    // 做多：第一層最接近 entry，後續價格依序往下（較低）
                    price = entryPrice - i * stepPct * entryPrice;
                }
                else
                {
                    // 做空：第一層最接近 entry，後續價格依序往上（較高）
                    price = entryPrice + i * stepPct * entryPrice;
                }
                long qty = finalQtys[i];
                if (qty > 0) outList.Add((price, qty));
            }
            return outList;
        }

        /// <summary>
        /// 把 qty 向下對齊成 lotSize 的倍數
        /// </summary>
        private static long NormalizeToLot(long qty, int lotSize)
        {
            if (lotSize <= 0) return 0;
            if (qty <= 0) return 0;
            return (qty / lotSize) * lotSize;
        }
    }
}