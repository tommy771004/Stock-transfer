# LLMAgentTrader 🤖📈

> **AI 驅動的量化交易分析平台**
> 結合 LLM 多代理人辯論框架、技術指標引擎與即時財經資料，輔助台股/美股投資決策。

> ⚠️ **免責聲明**：本軟體僅供學術研究與教育用途，不構成任何投資建議。股票投資有風險，請自行承擔交易結果。

---

## ✨ 主要功能

| 功能 | 說明 |
|---|---|
| 📊 技術指標分析 | SMA/EMA/RSI/MACD/ATR/布林通道/KD/DMI，一鍵計算 |
| 🤖 AI 多代理人辯論 | 多頭/空頭代理人互相辯論，產出 Buy/Sell/Hold 建議 |
| 🔍 股票篩選器 | 支援 SMA 多頭排列、RSI 超賣、MACD 黃金交叉等條件組合篩選 |
| 📐 回測引擎 | 含 Sharpe/Sortino/Kelly 等績效指標 |
| 💰 倉位管理 | Kelly Fraction + Half-Kelly 縮減 + 金字塔加碼計畫 |
| 🏛️ 三大法人資料 | 即時抓取外資/投信/自營商買賣超（台股） |
| 📉 融資融券監控 | TWSE/TPEX 融資餘額即時查詢 |
| 🌐 板塊輪動快照 | 11 大板塊 ETF 當日強弱排行 |

---

## 🚀 安裝與執行

### 環境需求
- **Windows 10/11** (64-bit)
- **.NET 8.0 SDK**（[下載](https://dotnet.microsoft.com/download/dotnet/8.0)）
- Visual Studio 2022 或 VS Code（含 C# 擴充）

### 執行步驟
```bash
git clone https://github.com/tommy771004/Stock-transfer.git
cd Stock-transfer
dotnet build LLMAgentTrader.sln
dotnet run --project LLMAgentTrader
```

---

## 🔑 API Key 設定

在程式內「設定」分頁填入，金鑰以 Windows DPAPI 加密儲存於本機：

| 服務 | 用途 | 取得網址 |
|---|---|---|
| OpenRouter | GPT-4o / Claude / Llama 等 LLM 路由 | https://openrouter.ai/keys |
| Google AI Studio | Gemini 2.0 Flash / 1.5 Pro 直連 | https://aistudio.google.com/apikey |
| Alpha Vantage | 美股基本面資料（PE/殖利率 Fallback） | https://www.alphavantage.co/support/#api-key |

---

## 🛠️ 專案結構

```
LLMAgentTrader/
├── Engine.cs              # 技術指標/回測/AI 代理引擎（核心類別）
├── Engines/               # 拆分出的獨立引擎模組
│   ├── SupportResistanceEngine.cs
│   ├── DivergenceDetector.cs
│   ├── FibonacciEngine.cs
│   ├── MultiTimeframeEngine.cs
│   ├── ScreenerEngine.cs
│   ├── SectorCorrelationEngine.cs
│   ├── VolatilityAdaptiveEngine.cs
│   └── PerformanceAttributionEngine.cs
├── Services.cs            # HTTP/LLM/Yahoo/TWSE 資料服務
├── LLMAgentTrader.cs      # 主視窗 UI
├── AutoTradeForm.cs       # 自動交易表單
├── AIAutoTradeModule.cs   # AI 自動交易模組
├── Models.cs              # 資料模型
├── PositionSizingService.cs # 倉位計算服務
├── Controls.cs            # 自訂控制項
└── UITheme.cs             # UI 主題設定
```

---

## 📜 License

MIT License