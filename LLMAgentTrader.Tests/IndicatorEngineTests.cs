using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using LLMAgentTrader;

namespace LLMAgentTrader.Tests
{
    [TestClass]
    public class IndicatorEngineTests
    {
        private static List<MarketData> MakeFlat(int count, double price = 100.0)
            => Enumerable.Range(0, count).Select(_ => new MarketData
            {
                Open = price, High = price + 1, Low = price - 1,
                Close = price, Volume = 1000
            }).ToList();

        [TestMethod]
        public void SMA5_ShouldEqualClose_WhenAllPricesIdentical()
        {
            var data = MakeFlat(30);
            IndicatorEngine.CalculateAll(data);
            Assert.AreEqual(100.0, data.Last().SMA5, 0.001,
                "當所有收盤價相同時，SMA5 應等於收盤價");
        }

        [TestMethod]
        public void SMA20_ShouldBeZero_WhenLessThan20Bars()
        {
            var data = MakeFlat(15);
            // CalculateAll 需要 >= 30 根，此情境下方法直接 return，SMA20 應為預設 0
            IndicatorEngine.CalculateAll(data);
            Assert.AreEqual(0.0, data.Last().SMA20, 0.001,
                "資料不足時 SMA20 應保持 0");
        }

        [TestMethod]
        public void SMA60_ShouldEqualClose_WhenAllPricesIdentical()
        {
            var data = MakeFlat(80);
            IndicatorEngine.CalculateAll(data);
            Assert.AreEqual(100.0, data.Last().SMA60, 0.001,
                "當所有收盤價相同時，SMA60 應等於收盤價");
        }

        [TestMethod]
        public void RSI_ShouldBe100_WhenOnlyGains()
        {
            var data = Enumerable.Range(0, 30).Select(i => new MarketData
            {
                Open  = 100 + i,
                High  = 100 + i + 1,
                Low   = 100 + i - 0.5,
                Close = 100 + i + 1,
                Volume = 1000
            }).ToList();
            IndicatorEngine.CalculateAll(data);
            Assert.IsTrue(data.Last().RSI > 90,
                $"純漲趨勢 RSI 應接近 100，實際={data.Last().RSI:F1}");
        }

        [TestMethod]
        public void Bias20_ShouldBeZero_WhenAllPricesIdentical()
        {
            var data = MakeFlat(50);
            IndicatorEngine.CalculateAll(data);
            Assert.AreEqual(0.0, data.Last().Bias20, 0.001,
                "當所有收盤價相同時，乖離率應為 0");
        }

        [TestMethod]
        public void SMA5_ShouldBeCorrect_AfterRollingUpdate()
        {
            // 前30根都是100，最後1根改為200 → SMA5 應為 (100*4 + 200)/5 = 120
            var data = MakeFlat(31);
            data[30] = new MarketData { Open = 200, High = 201, Low = 199, Close = 200, Volume = 1000 };
            IndicatorEngine.CalculateAll(data);
            Assert.AreEqual(120.0, data.Last().SMA5, 0.001,
                "滾動窗口 SMA5 計算應正確");
        }
    }
}
