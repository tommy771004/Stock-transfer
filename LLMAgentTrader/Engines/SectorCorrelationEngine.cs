using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LLMAgentTrader
{
    // ════════════════════════════════════════════════════════════════════════════
    //  ④ SectorCorrelationEngine
    //  板塊相關係數矩陣 + 領先/滯後偵測 + 相對強度排名
    //  輸入：各板塊近期歷史資料（由 SectorRotationService 的快照補充）
    // ════════════════════════════════════════════════════════════════════════════
    public static class SectorCorrelationEngine
    {
        // ── 相關係數矩陣 ─────────────────────────────────────────────────────
        /// <summary>
        /// 計算 n 個板塊之間的日報酬相關係數矩陣（Pearson）。
        /// sectorReturns[i] = 板塊 i 的日報酬序列（已對齊日期）
        /// </summary>
        public static double[,] CalcCorrelationMatrix(List<double[]> sectorReturns)
        {
            int n = sectorReturns.Count;
            var mat = new double[n, n];
            for (int i = 0; i < n; i++)
                for (int j = 0; j < n; j++)
                    mat[i, j] = i == j ? 1.0 : PearsonCorrelation(sectorReturns[i], sectorReturns[j]);
            return mat;
        }

        private static double PearsonCorrelation(double[] x, double[] y)
        {
            int len = Math.Min(x.Length, y.Length);
            if (len < 3) return 0;
            double meanX = x.Take(len).Average();
            double meanY = y.Take(len).Average();
            double cov = 0, varX = 0, varY = 0;
            for (int i = 0; i < len; i++)
            {
                double dx = x[i] - meanX, dy = y[i] - meanY;
                cov += dx * dy; varX += dx * dx; varY += dy * dy;
            }
            double denom = Math.Sqrt(varX * varY);
            return denom > 0 ? cov / denom : 0;
        }

        // ── 板塊相對強度（RS）排名 ────────────────────────────────────────────
        /// <summary>
        /// RS = 板塊 20 日報酬 / SPY 20 日報酬。
        /// snapshots：(ticker, name, emoji, changePct, price) 的快照
        /// spyReturn：SPY 同期報酬（傳 0 時降級為純漲幅排名）
        /// </summary>
        public static SectorCorrelationResult BuildResult(
            List<(string Ticker, string Name, string Emoji, double Return20d, double Price)> sectors,
            double spyReturn20d,
            List<double[]> returnSeries = null)
        {
            var result = new SectorCorrelationResult();

            // 相對強度排名
            var ranking = sectors.Select(s =>
            {
                double rs = spyReturn20d != 0 ? s.Return20d / Math.Abs(spyReturn20d) : s.Return20d;
                return (s.Ticker, s.Name , rs, s.Return20d);
            }).OrderByDescending(x => x.rs).ToList();
            result.Ranking = ranking;
            result.SectorNames = sectors.Select(s => s.Name).ToArray();

            // 領先/滯後
            result.LeadingSector = ranking.FirstOrDefault().Name ?? "";
            result.LaggingSector = ranking.LastOrDefault().Name ?? "";

            // 相關係數矩陣（若有資料序列）
            if (returnSeries?.Count >= 2)
            {
                result.Matrix = CalcCorrelationMatrix(returnSeries);

                // 輪動信號：找強弱差距大且最近反轉的板塊對
                var topRS = ranking.Take(3).Select(r => r.Ticker).ToHashSet();
                var bottomRS = ranking.TakeLast(3).Select(r => r.Ticker).ToHashSet();
                var signals = new List<(string, string, double)>();

                foreach (var top in ranking.Take(2))
                    foreach (var bot in ranking.TakeLast(2))
                    {
                        double gap = top.rs - bot.rs;
                        if (gap > 0.15)   // 強弱差距 > 15%
                        {
                            double conf = Math.Min(1.0, gap * 3);   // 最高 100%
                            signals.Add((bot.Name, top.Name, conf));
                        }
                    }
                result.RotationSignals = signals;
            }

            return result;
        }

        // ── 生成 AI Prompt 板塊情境 ──────────────────────────────────────────
        public static string ToPromptContext(SectorCorrelationResult r)
        {
            if (r == null || r.Ranking == null || r.Ranking.Count == 0) return "";
            var sb = new StringBuilder();
            sb.AppendLine("【板塊輪動情境】");
            sb.Append("強勢板塊：");
            sb.AppendLine(string.Join("  ", r.Ranking.Take(3).Select(x => $"{x.Name}(RS={x.RS:+0.00;-0.00})")));
            sb.Append("弱勢板塊：");
            sb.AppendLine(string.Join("  ", r.Ranking.TakeLast(3).Reverse().Select(x => $"{x.Name}(RS={x.RS:+0.00;-0.00})")));
            if (r.RotationSignals?.Count > 0)
            {
                sb.AppendLine("輪動訊號：");
                foreach (var (from, to, conf) in r.RotationSignals.Take(2))
                    sb.AppendLine($"  資金可能從 {from} → {to}（信心 {conf:P0}）");
            }
            return sb.ToString();
        }
    }
}
