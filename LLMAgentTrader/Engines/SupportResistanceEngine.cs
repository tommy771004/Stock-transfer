using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LLMAgentTrader
{
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
}
