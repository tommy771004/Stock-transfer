using System;

namespace LLMAgentTrader
{
    /// <summary>
    /// 回測結果容器
    /// </summary>
    public class Result
    {
        public double TotalReturn { get; set; }
        public double MaxDrawdown { get; set; }
        public double WinRate { get; set; }
        public int TradeCount { get; set; }
        public double SharpeRatio { get; set; }
        public double SortinoRatio { get; set; }
        public double KellyFraction { get; set; }
        public double AvgWinPct { get; set; }
        public double AvgLossPct { get; set; }
    }
}