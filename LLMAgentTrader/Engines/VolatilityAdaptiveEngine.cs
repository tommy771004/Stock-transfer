using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LLMAgentTrader
{
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
}
