using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LLMAgentTrader
{
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
}
