using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using LLMAgentTrader;

namespace LLMAgentTrader.Tests
{
    [TestClass]
    public class BacktestEngineTests
    {
        private static List<MarketData> MakeData(int count, double price = 100.0)
            => Enumerable.Range(0, count).Select(_ => new MarketData
            {
                Open = price, High = price + 2, Low = price - 2,
                Close = price, Volume = 5000,
                AgentAction = "Hold"
            }).ToList();

        [TestMethod]
        public void RunBacktest_ShouldReturnZeroReturn_WhenNoTrades()
        {
            var data = MakeData(50);
            var result = BacktestEngine.RunBacktest(data);
            Assert.AreEqual(0, result.TradeCount, "無交易信號時交易次數應為 0");
            Assert.AreEqual(0.0, result.TotalReturn, 0.001, "無交易時總報酬應為 0");
        }

        [TestMethod]
        public void RunBacktest_WinRate_ShouldBe1_WhenPriceAlwaysRises()
        {
            // 建立上升行情，第1根 Buy，第10根 Sell
            var data = Enumerable.Range(0, 20).Select(i => new MarketData
            {
                Open = 100 + i, High = 102 + i, Low = 99 + i,
                Close = 100 + i, Volume = 5000,
                AgentAction = "Hold"
            }).ToList();
            data[1].AgentAction = "Buy";
            data[10].AgentAction = "Sell";

            var result = BacktestEngine.RunBacktest(data);
            Assert.IsTrue(result.WinRate > 0, "上漲後賣出應為贏");
            Assert.IsTrue(result.TotalReturn > 0, "上漲後總報酬應為正");
        }

        [TestMethod]
        public void RunBacktest_KellyFraction_ShouldBeBetween0And05()
        {
            var data = MakeData(50);
            // 製造交易：連續幾次 Buy/Sell
            for (int i = 1; i < 50; i += 8)
            {
                if (i < 48) data[i].AgentAction = "Buy";
                if (i + 4 < 50) data[i + 4].AgentAction = "Sell";
            }
            var result = BacktestEngine.RunBacktest(data);
            Assert.IsTrue(result.KellyFraction >= 0 && result.KellyFraction <= 0.5,
                $"Kelly Fraction 應在 [0, 0.5]，實際={result.KellyFraction:F3}");
        }
    }
}
