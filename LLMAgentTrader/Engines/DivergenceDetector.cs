using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LLMAgentTrader
{
    // ────────────────────────────────────────────────────────────────────────────
    //  背離偵測引擎 (RSI 正/負背離、MACD 底/頂背離)
    //  原理：在一段回溯窗口內比較價格高/低點與指標高/低點的方向差異
    // ────────────────────────────────────────────────────────────────────────────
    public static class DivergenceDetector
    {
        private const int Lookback = 20;   // 回溯窗口（根K棒）
        private const int MinSwing = 5;    // 極點之間最少間隔根數

        public static void Detect(List<MarketData> list)
        {
            for (int i = Lookback + MinSwing; i < list.Count; i++)
            {
                var win = list.Skip(i - Lookback).Take(Lookback + 1).ToList();
                // ── RSI 背離 ─────────────────────────────────────────────────
                list[i].RSI_Divergence = DetectRsiDiv(win);
                // ── MACD 柱 背離 ─────────────────────────────────────────────
                list[i].MACD_Divergence = DetectMacdDiv(win);
            }
        }

        private static string DetectRsiDiv(List<MarketData> win)
        {
            // Bug #7 fix: 初始化為 -1，確保「未找到極值」時條件 lo2Idx > lo1Idx 不會誤判
            // 正背離：價格創低 但 RSI 未創低 → 多頭訊號
            int lo1Idx = -1, lo2Idx = -1;
            double lo1 = double.MaxValue, lo2 = double.MaxValue;
            for (int j = 1; j < win.Count - MinSwing; j++)
                if (win[j].Low < lo1) { lo1 = win[j].Low; lo1Idx = j; }
            if (lo1Idx >= 0)
            {
                for (int j = lo1Idx + MinSwing; j < win.Count; j++)
                    if (win[j].Low < lo2) { lo2 = win[j].Low; lo2Idx = j; }
                if (lo2Idx > lo1Idx && lo2 < lo1 && win[lo2Idx].RSI > win[lo1Idx].RSI && win[lo1Idx].RSI > 0)
                    return "正背離";
            }

            // 負背離：價格創高 但 RSI 未創高 → 空頭訊號
            int hi1Idx = -1, hi2Idx = -1;
            double hi1 = double.MinValue, hi2 = double.MinValue;
            for (int j = 1; j < win.Count - MinSwing; j++)
                if (win[j].High > hi1) { hi1 = win[j].High; hi1Idx = j; }
            if (hi1Idx >= 0)
            {
                for (int j = hi1Idx + MinSwing; j < win.Count; j++)
                    if (win[j].High > hi2) { hi2 = win[j].High; hi2Idx = j; }
                if (hi2Idx > hi1Idx && hi2 > hi1 && win[hi2Idx].RSI < win[hi1Idx].RSI && win[hi1Idx].RSI > 0)
                    return "負背離";
            }

            return "-";
        }

        private static string DetectMacdDiv(List<MarketData> win)
        {
            // Bug #7 fix: 初始化為 -1，確保「未找到極值」時條件不誤判
            // 底背離：價格低點更低 但 MACD 柱低點更高 → 多頭
            int lo1Idx = -1, lo2Idx = -1;
            double lo1 = double.MaxValue, lo2 = double.MaxValue;
            for (int j = 1; j < win.Count - MinSwing; j++)
                if (win[j].Low < lo1) { lo1 = win[j].Low; lo1Idx = j; }
            if (lo1Idx >= 0)
            {
                for (int j = lo1Idx + MinSwing; j < win.Count; j++)
                    if (win[j].Low < lo2) { lo2 = win[j].Low; lo2Idx = j; }
                if (lo2Idx > lo1Idx && lo2 < lo1 && win[lo2Idx].MACD_Hist > win[lo1Idx].MACD_Hist)
                    return "底背離";
            }

            // 頂背離：價格高點更高 但 MACD 柱高點更低 → 空頭
            int hi1Idx = -1, hi2Idx = -1;
            double hi1 = double.MinValue, hi2 = double.MinValue;
            for (int j = 1; j < win.Count - MinSwing; j++)
                if (win[j].High > hi1) { hi1 = win[j].High; hi1Idx = j; }
            if (hi1Idx >= 0)
            {
                for (int j = hi1Idx + MinSwing; j < win.Count; j++)
                    if (win[j].High > hi2) { hi2 = win[j].High; hi2Idx = j; }
                if (hi2Idx > hi1Idx && hi2 > hi1 && win[hi2Idx].MACD_Hist < win[hi1Idx].MACD_Hist)
                    return "頂背離";
            }

            return "-";
        }
    }
}
