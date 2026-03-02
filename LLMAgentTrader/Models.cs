using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace LLMAgentTrader
{
    public class MarketData
    {
        public DateTime Date { get; set; }
        public double Open { get; set; }
        public double Close { get; set; }
        public double High { get; set; }
        public double Low { get; set; }
        public double Volume { get; set; }

        public double RSI { get; set; }
        public double MACD { get; set; }
        public double MACD_Signal { get; set; }
        public double MACD_Hist { get; set; }
        public double EMA_50 { get; set; }
        public double EMA_200 { get; set; }
        public double BB_Upper { get; set; }
        public double BB_Lower { get; set; }
        public double BB_Middle { get; set; }
        public double BB_Width { get; set; }
        public double ATR { get; set; }
        public double SupportLevel { get; set; }
        public double ResistanceLevel { get; set; }
        public double VWAP { get; set; }

        // ── 新增指標 (來自學習資源：KD/Stochastic、DMI、SMA、乖離率) ─────────
        /// <summary>KD 隨機指標 %K (Fast Stochastic, 9期)</summary>
        public double KD_K { get; set; }
        /// <summary>KD 隨機指標 %D (Slow Stochastic, 3期平滑)</summary>
        public double KD_D { get; set; }
        /// <summary>DMI 多向指標 +DI (14期)</summary>
        public double DMI_Plus { get; set; }
        /// <summary>DMI 空向指標 -DI (14期)</summary>
        public double DMI_Minus { get; set; }
        /// <summary>ADX 趨勢強度 (14期)</summary>
        public double DMI_ADX { get; set; }
        /// <summary>SMA 5日均線</summary>
        public double SMA5 { get; set; }
        /// <summary>SMA 10日均線</summary>
        public double SMA10 { get; set; }
        /// <summary>SMA 20日均線</summary>
        public double SMA20 { get; set; }
        /// <summary>SMA 60日均線</summary>
        public double SMA60 { get; set; }
        /// <summary>乖離率 Bias：(Close - SMA20) / SMA20 × 100</summary>
        public double Bias20 { get; set; }
        /// <summary>布林通道是否緊縮（BB_Width &lt; 3.5%）</summary>
        public bool BB_Squeeze { get; set; }
        /// <summary>RSI 背離信號：正背離/負背離/-</summary>
        public string RSI_Divergence { get; set; } = "-";
        /// <summary>MACD 背離信號：底背離/頂背離/-</summary>
        public string MACD_Divergence { get; set; } = "-";
        /// <summary>複合 K 線型態（三根K棒組合：晨星、夜星、紅三兵、黑三兵等）</summary>
        public string Pattern2 { get; set; } = "-";

        public string Pattern { get; set; } = "-";
        public string AgentAction { get; set; } = "Hold";
        public string AgentReasoning { get; set; } = "";
    }

    public class TwseTick
    {
        public string Time { get; set; }
        public double Price { get; set; }
        public double YesterdayClose { get; set; }
        public string CompanyName { get; set; }
        public double Volume { get; set; }
        public double TradeVolume { get; set; }
        public double Open { get; set; }
        public double High { get; set; }
        public double Low { get; set; }
        public List<double> AskPrices { get; set; } = new List<double>();
        public List<int> AskVolumes { get; set; } = new List<int>();
        public List<double> BidPrices { get; set; } = new List<double>();
        public List<int> BidVolumes { get; set; } = new List<int>();
    }

    public class LlmResult
    {
        [JsonPropertyName("Date")] public string Date { get; set; }
        [JsonPropertyName("Action")] public string Action { get; set; }
        [JsonPropertyName("Reasoning")] public string Reasoning { get; set; }
    }

    public class FeedbackRecord
    {
        public DateTime Date { get; set; }
        public string Ticker { get; set; }
        public string Action { get; set; }
        public double EntryPrice { get; set; }
        public double RSI { get; set; }
        public string Outcome { get; set; }
        public string Lesson { get; set; }
    }

    // ── Fibonacci 層級 ────────────────────────────────────────────────────────
    public class FibLevel
    {
        public double Ratio { get; set; }
        public double Price { get; set; }
        public string Label { get; set; }
    }

    // ── 多時間框架信號 ────────────────────────────────────────────────────────
    public class MultiTimeframeSignal
    {
        public string Ticker { get; set; }
        public double Weekly_RSI { get; set; }
        public double Weekly_MACD_Hist { get; set; }
        public string Weekly_Trend { get; set; } = "-";
        public string Weekly_Pattern { get; set; } = "-";
        public double Weekly_KD_K { get; set; }
        public double Weekly_ADX { get; set; }
        public string Weekly_RSI_Div { get; set; } = "-";
        public double Daily_RSI { get; set; }
        public double Daily_MACD_Hist { get; set; }
        public string Daily_Trend { get; set; } = "-";
        public string Daily_Pattern { get; set; } = "-";
        public double Daily_KD_K { get; set; }
        public double Daily_ADX { get; set; }
        public string Daily_RSI_Div { get; set; } = "-";
        public string Daily_MACD_Div { get; set; } = "-";
        public string Daily_Pattern2 { get; set; } = "-";
        public double Hourly_RSI { get; set; }
        public double Hourly_MACD_Hist { get; set; }
        public string Hourly_Trend { get; set; } = "-";
        public string Hourly_Pattern { get; set; } = "-";
        public double Hourly_KD_K { get; set; }
        public int AlignmentScore { get; set; }
        public string AlignmentSummary { get; set; } = "";
    }

    // ── 交易日誌條目 ──────────────────────────────────────────────────────────
    public class TradeJournalEntry
    {
        public int Id { get; set; }
        public DateTime TradeDate { get; set; }
        public string Ticker { get; set; }
        public string Direction { get; set; }
        public double EntryPrice { get; set; }
        public double ExitPrice { get; set; }
        public double Quantity { get; set; }
        public double StopLossPrice { get; set; }   // 停損價 (0 = 未設)
        public double TargetPrice { get; set; }     // 目標價 (0 = 未設)
        public string AiSuggestion { get; set; } = "-";
        public string MyDecision { get; set; } = "N/A";
        public string Notes { get; set; } = "";

        public double PnL => (EntryPrice > 0 && ExitPrice > 0)
            ? (ExitPrice - EntryPrice) * Quantity * (Direction == "Buy" ? 1 : -1)
            : 0;
        public double ReturnPct => (EntryPrice > 0 && ExitPrice > 0)
            ? (ExitPrice - EntryPrice) / EntryPrice * (Direction == "Buy" ? 1 : -1)
            : 0;
        public string Status => ExitPrice > 0 ? (PnL >= 0 ? "獲利" : "虧損") : "持倉中";

        // 風報比 (Target - Entry) / (Entry - StopLoss)，做多方向
        public double RiskRewardRatio => (EntryPrice > 0 && StopLossPrice > 0 && TargetPrice > 0 && EntryPrice != StopLossPrice)
            ? Math.Abs(TargetPrice - EntryPrice) / Math.Abs(EntryPrice - StopLossPrice)
            : 0;
    }

    // ── ATR / Fibonacci 停損停利建議 ─────────────────────────────────────────
    public class RiskRewardSuggestion
    {
        public double EntryPrice { get; set; }
        public string Method { get; set; }       // "ATR" 或 "Fibonacci"
        public double StopLoss { get; set; }
        public double Target1 { get; set; }      // 1R (風報比 1:1)
        public double Target2 { get; set; }      // 2R (風報比 1:2)
        public double Target3 { get; set; }      // 3R (風報比 1:3)
        public double AtrValue { get; set; }
        public double RiskPct { get; set; }      // 停損幅度 %
        public string FibStopLabel { get; set; } // Fib 方法下停損依據的層級名
        public string FibTargetLabel { get; set; }
        public string Summary { get; set; }
    }

    // ── 市場整體情緒 ─────────────────────────────────────────────────────────
    public class MarketSentiment
    {
        public DateTime FetchTime { get; set; }
        public double VIX { get; set; }
        public double VIXChange { get; set; }
        public string VIXLevel { get; set; }     // "低波動 (<15)", "正常 (15-25)", "恐慌 (25-35)", "極度恐慌 (>35)"
        public double SP500ChangePct { get; set; }
        public double NasdaqChangePct { get; set; }
        public double FearGreedScore { get; set; }  // 0-100 合成分數
        public string FearGreedLabel { get; set; }  // "極度恐懼 / 恐懼 / 中立 / 貪婪 / 極度貪婪"
        public string AiPositionAdvice { get; set; } // AI 倉位建議 (保守/中性/積極)
        public string Summary { get; set; }
        public bool IsStale => (DateTime.Now - FetchTime).TotalMinutes > 30;
    }

    // ── 止盈止損警報追蹤 ──────────────────────────────────────────────────────
    public class StopLossAlert
    {
        public int JournalId { get; set; }
        public string Ticker { get; set; }
        public string Direction { get; set; }
        public double EntryPrice { get; set; }
        public double StopLossPrice { get; set; }
        public double TargetPrice { get; set; }
        public bool IsActive { get; set; } = true;
        public bool StopTriggered { get; set; }
        public bool TargetTriggered { get; set; }
        public DateTime? TriggeredAt { get; set; }
        public string TriggeredType { get; set; } = ""; // "StopLoss" 或 "Target"
    }

    // ── 篩選條件 ──────────────────────────────────────────────────────────────
    public class ScreenerCriteria
    {
        public double RSI_Min { get; set; } = 0;
        public double RSI_Max { get; set; } = 100;
        public bool MACD_Positive { get; set; } = false;
        public bool MACD_CrossUp { get; set; } = false;
        public bool Above_EMA50 { get; set; } = false;
        public bool Above_EMA200 { get; set; } = false;
        public bool BB_Breakout { get; set; } = false;
        public bool BB_Oversold { get; set; } = false;
        public double Volume_Min_Ratio { get; set; } = 0;
        // ── 新增篩選條件 ────────────────────────────────────────────────────
        /// <summary>KD %K 超賣（K &lt; 20）</summary>
        public bool KD_Oversold { get; set; } = false;
        /// <summary>KD %K 超買（K &gt; 80）</summary>
        public bool KD_Overbought { get; set; } = false;
        /// <summary>KD 黃金交叉（K 由下向上穿越 D）</summary>
        public bool KD_CrossUp { get; set; } = false;
        /// <summary>DMI 多頭：+DI > -DI 且 ADX > 25（強趨勢多頭）</summary>
        public bool DMI_Bullish { get; set; } = false;
        /// <summary>布林通道緊縮（BB_Width &lt; 3.5%，蓄勢待發）</summary>
        public bool BB_Squeeze { get; set; } = false;
        /// <summary>RSI 正背離（多頭背離信號）</summary>
        public bool RSI_BullDiv { get; set; } = false;
        /// <summary>MACD 底背離（多頭背離信號）</summary>
        public bool MACD_BullDiv { get; set; } = false;
        /// <summary>SMA 多頭排列 (SMA5 > SMA10 > SMA20 > SMA60)</summary>
        public bool SMA_BullArrange { get; set; } = false;
    }

    // ── 篩選結果 ─────────────────────────────────────────────────────────────
    public class ScreenerResult
    {
        public string Ticker { get; set; }
        public string CompanyName { get; set; }
        public double Close { get; set; }
        public double RSI { get; set; }
        public double MACD_Hist { get; set; }
        public double EMA50 { get; set; }
        public double BB_Width { get; set; }
        public string Trend { get; set; }
        public string Pattern { get; set; }
        public int MatchScore { get; set; }
        public List<string> MatchedRules { get; set; } = new List<string>();
        // ── 新增篩選結果欄位 ────────────────────────────────────────────────
        public double KD_K { get; set; }
        public double KD_D { get; set; }
        public double DMI_ADX { get; set; }
        public double DMI_Plus { get; set; }
        public double DMI_Minus { get; set; }
        public double Bias20 { get; set; }
        public string RSI_Divergence { get; set; } = "-";
        public string MACD_Divergence { get; set; } = "-";
        public string Pattern2 { get; set; } = "-";
    }

    // ── ETF 資訊模型 ─────────────────────────────────────────────────────────
    public class EtfInfo
    {
        public string Ticker { get; set; } = "";
        public string Name { get; set; } = "";
        public bool IsEtf { get; set; } = false;
        public string Category { get; set; } = "";      // "股票型" / "債券型" / "商品型" etc.
        public string TrackIndex { get; set; } = "";      // 追蹤指數名稱
        public double TotalAssets { get; set; } = 0;       // AUM (USD or TWD)
        public double ExpenseRatio { get; set; } = 0;       // 管理費 (0.03 = 0.03%)
        public double Nav { get; set; } = 0;       // 淨值 (TWD ETF)
        public double Premium { get; set; } = 0;       // 折溢價 % (市價 vs NAV)
        public double YtdReturn { get; set; } = 0;       // 今年以來報酬 %
        public double ThreeYrReturn { get; set; } = 0;     // 3年年化報酬 %
        public double DividendYield { get; set; } = 0;     // 殖利率
        public List<(string Name, double Pct)> TopHoldings { get; set; } = new();
        public string AssetClass { get; set; } = "";      // "equity" / "bond" / "commodity" / "balanced"
        public string Source { get; set; } = "";      // 資料來源
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  ② PositionSizingEngine — 相關模型
    // ════════════════════════════════════════════════════════════════════════════

    /// <summary>Kelly 公式計算後的倉位建議</summary>
    public class PositionSizingResult
    {
        public double Capital { get; set; }   // 可用資金
        public double KellyFraction { get; set; }   // Kelly 比例 (0~0.5)
        public double RiskAmount { get; set; }   // 最大風險金額 = Capital * Kelly
        public double EntryPrice { get; set; }
        public double StopLossPrice { get; set; }
        public double StopLossPct { get; set; }   // 停損幅度
        public double RecommendedQty { get; set; }   // 建議買入張數/股數
        public double MaxPositionValue { get; set; }   // 最大持倉市值
        public string Summary { get; set; }
    }

    /// <summary>金字塔加碼 / 分批進場的單一層級</summary>
    public class PyramidLevel
    {
        public int Batch { get; set; }     // 批次編號（1 = 首批）
        public double Price { get; set; }     // 建議進場價
        public double QtyPct { get; set; }     // 佔總倉位 % (0~1)
        public double Qty { get; set; }     // 建議股數/張數
        public string Label { get; set; }     // 說明文字
    }

    /// <summary>分批出場（分層獲利了結）</summary>
    public class ExitLevel
    {
        public int Layer { get; set; }     // 出場層 1/2/3
        public double Price { get; set; }     // 建議出場價
        public double QtyPct { get; set; }     // 出場比例
        public string Rationale { get; set; }     // 依據（1R/2R/3R 或 Fib）
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  ③ SentimentAwareTradingEngine — 相關模型
    // ════════════════════════════════════════════════════════════════════════════

    /// <summary>根據市場情緒計算的風險上限</summary>
    public class SentimentRiskLimit
    {
        public double MaxPositionPct { get; set; }   // 最大持倉比例（0~1）
        public double MaxDrawdownPct { get; set; }   // 可接受最大回撤
        public string Regime { get; set; }   // "恐慌" / "正常" / "貪婪" / "極端"
        public string PromptModifier { get; set; }   // 插入 AI Prompt 的額外風控指示
        public bool VixJumpAlert { get; set; }   // VIX 跳升警報
        public double VixAcceleration { get; set; }   // VIX 加速度（今日變化 / 5日平均變化）
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  ④ SectorCorrelationEngine — 相關模型
    // ════════════════════════════════════════════════════════════════════════════

    /// <summary>板塊相關性分析結果</summary>
    public class SectorCorrelationResult
    {
        /// <summary>相關係數矩陣；索引對應 SectorRotationService.SectorEtfs</summary>
        public double[,] Matrix { get; set; }
        public string[] SectorNames { get; set; }
        /// <summary>板塊相對強度排名（高→低）</summary>
        public List<(string Ticker, string Name, double RS, double Return20d)> Ranking { get; set; } = new();
        /// <summary>偵測到的輪動訊號</summary>
        public List<(string From, string To, double Confidence)> RotationSignals { get; set; } = new();
        /// <summary>領先板塊（最近 5 日率先反彈/下跌者）</summary>
        public string LeadingSector { get; set; } = "";
        public string LaggingSector { get; set; } = "";
        public DateTime AnalysisTime { get; set; } = DateTime.Now;
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  ⑤ VolatilityAdaptiveEngine — 相關模型
    // ════════════════════════════════════════════════════════════════════════════

    /// <summary>波動率完整剖面</summary>
    public class VolatilityProfile
    {
        public double CurrentATR { get; set; }   // 當前 ATR (14日)
        public double AdaptiveATR { get; set; }   // VIX 調整後的 ATR
        public double VolPercentile { get; set; }   // 歷史波動率百分位 (0~100)
        public string VolRegime { get; set; }   // "低波動" / "正常" / "高波動" / "極端"
        public double AtrMultiplier { get; set; }   // 建議停損倍數
        public double SuggestedStop { get; set; }   // 建議停損距離 (價格單位)
        public string Summary { get; set; }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  ⑥ PerformanceAttributionEngine — 相關模型
    // ════════════════════════════════════════════════════════════════════════════

    /// <summary>績效分解結果</summary>
    public class PerformanceAttribution
    {
        public double SelectionEffect { get; set; }  // 標的選擇貢獻（alpha）
        public double TimingEffect { get; set; }  // 進場時機貢獻
        public double SizingEffect { get; set; }  // 倉位管理貢獻
        public double TotalEffect { get; set; }  // 合計
        public double AvgEntryQuality { get; set; }  // 平均進場品質評分 (0~100)
        public double WinRateByRR { get; set; }  // 風報比 ≥ 2 時的勝率
        /// <summary>持倉時間 → 平均報酬率（分 bucket）</summary>
        public Dictionary<string, double> HoldingPeriodBias { get; set; } = new();
        public string Summary { get; set; }
    }

    /// <summary>單筆交易的進場品質評分</summary>
    public class EntryQualityScore
    {
        public int TradeId { get; set; }
        public double Score { get; set; }   // 0~100
        public string Grade { get; set; }   // A/B/C/D
        public string RsiSignal { get; set; }
        public string MacdSignal { get; set; }
        public string PatternSignal { get; set; }
        public string Remark { get; set; }
    }


    // ════════════════════════════════════════════════════════════════════════════
    //  自動模擬交易模型
    // ════════════════════════════════════════════════════════════════════════════

    public class AutoTradeConfig
    {
        public string Ticker { get; set; } = "AAPL";
        public double DailyCapital { get; set; } = 100_000;  // 每日可用本金
        public int MaxBuyOrders { get; set; } = 3;         // 每日最大買單數
        public int MaxSellOrders { get; set; } = 3;         // 每日最大賣單數
        public double MaxPositionPct { get; set; } = 0.30;      // 單倉最大佔比
        public double StopLossPct { get; set; } = 0.05;      // 停損 %
        public double TakeProfitPct { get; set; } = 0.10;      // 停利 %
        public int PollIntervalSec { get; set; } = 30;        // 輪詢間隔（秒）
        public bool UseKellySize { get; set; } = true;      // 用 Kelly 計算倉位
        public string TradingStyle { get; set; } = "Swing";   // Day / Swing / Position
        public double MinSignalScore { get; set; } = 65;        // AI 信號最低分（0~100）
        public bool AutoJournalLog { get; set; } = true;      // 自動記入交易日誌
        public double TrailingStopPct { get; set; } = 0.0;   // 移動停損 %（0 = 關閉）
        public double MaxDailyLossPct { get; set; } = 0.0;   // 每日最大虧損 %（0 = 關閉）
    }

    public class SimulatedOrder
    {
        public int OrderId { get; set; }
        public DateTime OrderTime { get; set; }
        public string Ticker { get; set; }
        public string Side { get; set; }  // "Buy" / "Sell"
        public string OrderType { get; set; }  // "Market" / "Limit"
        public double LimitPrice { get; set; }  // 限價（0 = 市價）
        public double Qty { get; set; }
        public double FilledPrice { get; set; }
        public DateTime? FilledTime { get; set; }
        public string Status { get; set; }  // "Pending" / "Filled" / "Cancelled"
        public string AiReason { get; set; }  // AI 決策理由
        public double StopLoss { get; set; }
        public double TakeProfit { get; set; }
        public double RealizedPnL { get; set; }  // 成交後損益
    }

    // ── 三大法人籌碼資料（TWSE/TPEX 官方 API） ───────────────────────────────
    public class InstitutionalData
    {
        public string Ticker { get; set; } = "";
        public string Date { get; set; } = "";
        public string Source { get; set; } = "";
        /// <summary>外資買賣超（張，正=買超，負=賣超）</summary>
        public long ForeignNet { get; set; }
        /// <summary>投信買賣超（張）</summary>
        public long TrustNet { get; set; }
        /// <summary>自營商買賣超（張）</summary>
        public long DealerNet { get; set; }
        /// <summary>三大法人合計買賣超（張）</summary>
        public long TotalNet { get; set; }
    }

    // ── 融資融券資料（TWSE/TPEX 官方 API） ──────────────────────────────────
    public class MarginData
    {
        public string Ticker { get; set; } = "";
        public string Date { get; set; } = "";
        public string Source { get; set; } = "";
        /// <summary>當日融資買進（張）</summary>
        public long MarginBuy { get; set; }
        /// <summary>當日融資賣出（張）</summary>
        public long MarginSell { get; set; }
        /// <summary>融資餘額（張）</summary>
        public long MarginBal { get; set; }
        /// <summary>當日融券賣出（張）</summary>
        public long ShortSell { get; set; }
        /// <summary>當日融券買進（張）</summary>
        public long ShortBuy { get; set; }
        /// <summary>融券餘額（張）</summary>
        public long ShortBal { get; set; }
        /// <summary>券資比 = 融券餘額 / 融資餘額 × 100（%）；>20% 有軋空機會</summary>
        public double ShortRatio { get; set; }
    }

    public class AutoTradeSession
    {
        public DateTime StartTime { get; set; } = DateTime.Now;
        public DateTime EndTime { get; set; }
        public bool IsRunning { get; set; }
        public double InitialCapital { get; set; }
        public double CurrentCapital { get; set; }
        public double HoldingValue { get; set; }  // 持倉市值
        public double TotalPnL => CurrentCapital + HoldingValue - InitialCapital;
        public double TotalPnLPct => InitialCapital > 0 ? TotalPnL / InitialCapital : 0;
        public int TotalOrders { get; set; }
        public int FilledOrders { get; set; }
        public int WinCount { get; set; }
        public int LossCount { get; set; }
        public double WinRate => (WinCount + LossCount) > 0
                                          ? (double)WinCount / (WinCount + LossCount) : 0;
        public List<SimulatedOrder> Orders { get; set; } = new List<SimulatedOrder>();
        public List<string> Log { get; set; } = new List<string>();
    }
}