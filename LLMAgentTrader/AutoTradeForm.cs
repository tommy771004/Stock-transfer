using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LLMAgentTrader
{
    // ════════════════════════════════════════════════════════════════════════════
    //  AutoTradeForm  —  AI 模擬自動交易視窗
    //
    //  流程：
    //  1. 使用者設定參數（資金/股票/倉位限制等）
    //  2. 點擊「▶ 啟動自動交易」
    //  3. 每 N 秒：
    //     a. 抓取最新即時報價 + 技術指標
    //     b. 呼叫 LLM 做 Buy/Sell/Hold 決策（含信號評分 0~100）
    //     c. 信號分數 ≥ 設定閾值 → 掛單（模擬限價/市價）
    //     d. 確認成交（模擬撮合：次一根 K 棒開盤價成交）
    //     e. 觸發停損/停利 → 自動出場
    //     f. 結果自動寫入交易日誌（Excel）
    // ════════════════════════════════════════════════════════════════════════════
    public class AutoTradeForm : Form
    {
        // ── 設定參數 ──────────────────────────────────────────────────────────
        private AutoTradeConfig _config = new AutoTradeConfig();
        private AutoTradeSession _session = null;
        private CancellationTokenSource _tradeCts = null;
        private string _apiKey = "";
        private int _orderIdSeq = 1;
        // MTF 信號快取（每 5 次輪詢更新一次，避免頻繁網路請求）
        private double _mtfAlignmentScore = 50.0;   // 0-100，預設中性
        private int _mtfUpdateCounter = 0;

        // ── UI 元件 ───────────────────────────────────────────────────────────
        private TextBox txtTicker, txtApiKey;
        private NumericUpDown numCapital, numMaxBuy, numMaxSell;
        private NumericUpDown numStopLoss, numTakeProfit, numMinScore, numPollSec;
        private NumericUpDown numMaxPosPct, numTrailingStop, numMaxDailyLoss;
        private ComboBox cbStyle;
        private CheckBox chkKelly, chkAutoJournal;
        private Button btnStart, btnStop, btnClear;
        private RichTextBox txtLog;
        private DataGridView dgOrders;
        private Label lblStatus, lblCapital, lblPnL, lblWinRate, lblOrders;
        private Panel pnlStats;
        private System.Windows.Forms.Timer _uiTimer;

        public AutoTradeForm(string apiKey = "", double suggestedKelly = 0)
        {
            _apiKey = apiKey;
            InitUI();
            // 若主程式有回測 Kelly 建議，自動帶入倉位上限
            if (suggestedKelly > 0)
            {
                decimal suggestedPct = (decimal)Math.Round(Math.Min(suggestedKelly * 100, 45), 0);
                if (suggestedPct >= numMaxPosPct.Minimum && suggestedPct <= numMaxPosPct.Maximum)
                    numMaxPosPct.Value = suggestedPct;
                Log($"📊 已套用回測 Kelly 建議倉位 {suggestedKelly:P1} → {suggestedPct:F0}%（上限 45%）", Color.Gold);
            }
            _uiTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            _uiTimer.Tick += (s, e) => RefreshStats();
            _uiTimer.Start();
        }

        // ════════════════════════════════════════════════════════════════════════
        //  UI 建構
        // ════════════════════════════════════════════════════════════════════════
        private void InitUI()
        {
            Text = "🤖 AI 模擬自動交易";
            Size = new Size(1080, 720);
            MinimumSize = new Size(900, 600);
            BackColor = Color.FromArgb(14, 16, 24);
            ForeColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;

            var mainSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 5,
                SplitterDistance = 330,
                BackColor = Color.FromArgb(30, 32, 40)
            };
            Controls.Add(mainSplit);

            // ── 左側：參數設定 ────────────────────────────────────────────────
            BuildConfigPanel(mainSplit.Panel1);

            // ── 右側：執行監控 ────────────────────────────────────────────────
            BuildMonitorPanel(mainSplit.Panel2);
        }

        private void BuildConfigPanel(Panel parent)
        {
            var tlp = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 17,
                BackColor = Color.FromArgb(18, 20, 30),
                Padding = new Padding(12, 10, 12, 10),
                CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
            };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            parent.Controls.Add(tlp);

            int row = 0;

            // 標題
            var title = new Label
            {
                Text = "⚙️ 自動交易設定",
                Font = new Font("Segoe UI", 13F, FontStyle.Bold),
                ForeColor = Color.FromArgb(120, 200, 255),
                Dock = DockStyle.Fill,
                Height = 36,
                TextAlign = ContentAlignment.MiddleLeft
            };
            tlp.Controls.Add(title, 0, row); tlp.SetColumnSpan(title, 2); row++;

            // API Key（顯示但可覆寫）
            tlp.Controls.Add(MakeLabel("API Key"), 0, row);
            txtApiKey = MakeTextBox(_apiKey, true);
            tlp.Controls.Add(txtApiKey, 1, row); row++;

            // 股票代號
            tlp.Controls.Add(MakeLabel("股票代號"), 0, row);
            txtTicker = MakeTextBox("AAPL");
            tlp.Controls.Add(txtTicker, 1, row); row++;

            // 每日本金
            tlp.Controls.Add(MakeLabel("每日本金 ($)"), 0, row);
            numCapital = MakeNumeric(100_000, 10_000, 10_000_000, 50_000, 0, true);
            tlp.Controls.Add(numCapital, 1, row); row++;

            // 買/賣單上限
            tlp.Controls.Add(MakeLabel("最大買單數"), 0, row);
            numMaxBuy = MakeNumeric(3, 1, 20, 1, 0);
            tlp.Controls.Add(numMaxBuy, 1, row); row++;

            tlp.Controls.Add(MakeLabel("最大賣單數"), 0, row);
            numMaxSell = MakeNumeric(3, 1, 20, 1, 0);
            tlp.Controls.Add(numMaxSell, 1, row); row++;

            // 倉位上限
            tlp.Controls.Add(MakeLabel("單倉上限 (%)"), 0, row);
            numMaxPosPct = MakeNumeric(30, 5, 100, 5, 0);
            tlp.Controls.Add(numMaxPosPct, 1, row); row++;

            // 停損/停利
            tlp.Controls.Add(MakeLabel("停損 (%)"), 0, row);
            numStopLoss = MakeNumeric(5, 1, 50, 1, 1);
            tlp.Controls.Add(numStopLoss, 1, row); row++;

            tlp.Controls.Add(MakeLabel("停利 (%)"), 0, row);
            numTakeProfit = MakeNumeric(10, 1, 100, 1, 1);
            tlp.Controls.Add(numTakeProfit, 1, row); row++;

            // 移動停損
            tlp.Controls.Add(MakeLabel("移動停損 (%)"), 0, row);
            numTrailingStop = MakeNumeric(0, 0, 30, 1, 1);
            numTrailingStop.Width = 80;
            var pnlTS = new FlowLayoutPanel { Dock = DockStyle.Fill, BackColor = Color.Transparent, WrapContents = false };
            pnlTS.Controls.Add(numTrailingStop);
            pnlTS.Controls.Add(new Label { Text = "0 = 關閉", ForeColor = Color.Gray, Font = new Font("Segoe UI", 8F), AutoSize = true, Padding = new Padding(4, 8, 0, 0) });
            tlp.Controls.Add(pnlTS, 1, row); row++;

            // 每日最大虧損上限
            tlp.Controls.Add(MakeLabel("每日虧損上限(%)"), 0, row);
            numMaxDailyLoss = MakeNumeric(0, 0, 50, 1, 1);
            numMaxDailyLoss.Width = 80;
            var pnlDL = new FlowLayoutPanel { Dock = DockStyle.Fill, BackColor = Color.Transparent, WrapContents = false };
            pnlDL.Controls.Add(numMaxDailyLoss);
            pnlDL.Controls.Add(new Label { Text = "0 = 關閉", ForeColor = Color.Gray, Font = new Font("Segoe UI", 8F), AutoSize = true, Padding = new Padding(4, 8, 0, 0) });
            tlp.Controls.Add(pnlDL, 1, row); row++;

            // 輪詢間隔
            tlp.Controls.Add(MakeLabel("輪詢間隔 (秒)"), 0, row);
            numPollSec = MakeNumeric(30, 5, 300, 5, 0);
            tlp.Controls.Add(numPollSec, 1, row); row++;

            // AI 最低信號分
            tlp.Controls.Add(MakeLabel("AI 最低信號分"), 0, row);
            numMinScore = MakeNumeric(65, 0, 100, 5, 0);
            tlp.Controls.Add(numMinScore, 1, row); row++;

            // 交易風格
            tlp.Controls.Add(MakeLabel("交易風格"), 0, row);
            cbStyle = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(40, 42, 55),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10F),
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat
            };
            cbStyle.Items.AddRange(new object[] { "Day（當日沖銷）", "Swing（波段）", "Position（長線）" });
            cbStyle.SelectedIndex = 1;
            tlp.Controls.Add(cbStyle, 1, row); row++;

            // 選項
            chkKelly = new CheckBox
            {
                Text = "使用 Kelly 公式計算倉位",
                Checked = true,
                ForeColor = Color.LightGray,
                Font = new Font("Segoe UI", 9.5F),
                Dock = DockStyle.Fill,
                AutoSize = false,
                Height = 26
            };
            tlp.Controls.Add(chkKelly, 0, row); tlp.SetColumnSpan(chkKelly, 2); row++;

            chkAutoJournal = new CheckBox
            {
                Text = "自動寫入交易日誌（Excel）",
                Checked = true,
                ForeColor = Color.LightGray,
                Font = new Font("Segoe UI", 9.5F),
                Dock = DockStyle.Fill,
                AutoSize = false,
                Height = 26
            };
            tlp.Controls.Add(chkAutoJournal, 0, row); tlp.SetColumnSpan(chkAutoJournal, 2); row++;

            // 按鈕列
            var pnlBtns = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Padding = new Padding(0, 6, 0, 0)
            };
            btnStart = MakeButton("▶ 啟動自動交易", Color.FromArgb(0, 140, 80));
            btnStop = MakeButton("⏹ 停止", Color.FromArgb(140, 40, 40));
            btnClear = MakeButton("🗑 清除日誌", Color.FromArgb(50, 52, 65));
            btnStop.Enabled = false;
            btnStart.Click += BtnStart_Click;
            btnStop.Click += BtnStop_Click;
            btnClear.Click += (s, e) => { txtLog.Clear(); };
            pnlBtns.Controls.Add(btnStart);
            pnlBtns.Controls.Add(btnStop);
            pnlBtns.Controls.Add(btnClear);
            tlp.Controls.Add(pnlBtns, 0, row); tlp.SetColumnSpan(pnlBtns, 2);

            // 設定行高
            for (int i = 0; i < tlp.RowCount; i++)
                tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, i == 0 ? 40 : 38));
        }

        private void BuildMonitorPanel(Panel parent)
        {
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterWidth = 5,
                SplitterDistance = 200,
                BackColor = Color.FromArgb(22, 24, 34)
            };
            parent.Controls.Add(split);

            // ── 上方：統計 Dashboard ──────────────────────────────────────────
            pnlStats = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(18, 20, 30), Padding = new Padding(10, 8, 10, 8) };
            var statsTitle = new Label
            {
                Text = "📊 即時交易統計",
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                ForeColor = Color.FromArgb(100, 200, 255),
                Dock = DockStyle.Top,
                Height = 30
            };
            pnlStats.Controls.Add(statsTitle);

            var statsFlow = new FlowLayoutPanel { Dock = DockStyle.Fill, BackColor = Color.Transparent, WrapContents = false };
            lblCapital = MakeStat("💰 可用本金", "-");
            lblPnL = MakeStat("📈 累計損益", "-");
            lblWinRate = MakeStat("🎯 勝率", "-");
            lblOrders = MakeStat("📋 成交/總單", "-");
            lblStatus = new Label
            {
                Text = "⏸ 未啟動",
                ForeColor = Color.LightGray,
                Font = new Font("Segoe UI", 10F),
                AutoSize = false,
                Width = 350,
                Dock = DockStyle.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0)
            };
            statsFlow.Controls.Add(lblCapital);
            statsFlow.Controls.Add(lblPnL);
            statsFlow.Controls.Add(lblWinRate);
            statsFlow.Controls.Add(lblOrders);
            statsFlow.Controls.Add(lblStatus);
            pnlStats.Controls.Add(statsFlow);
            split.Panel1.Controls.Add(pnlStats);

            // ── 下方：訂單列表 + 執行日誌 ────────────────────────────────────
            var innerSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 480,
                SplitterWidth = 5,
                BackColor = Color.FromArgb(22, 24, 34)
            };
            split.Panel2.Controls.Add(innerSplit);

            // 訂單 Grid
            dgOrders = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.FromArgb(14, 16, 24),
                ForeColor = Color.White,
                GridColor = Color.FromArgb(40, 42, 55),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ReadOnly = true,
                Font = new Font("Consolas", 9F),
                EnableHeadersVisualStyles = false,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(30, 32, 45),
                    ForeColor = Color.LightGray,
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold)
                }
            };
            var cols = new (string Name, string Header, int Width)[]
            {
                ("OrderId",    "#",       42),
                ("Side",       "方向",    50),
                ("OrderType",  "類型",    55),
                ("LimitPrice", "掛單價",  70),
                ("FilledPrice","成交價",  70),
                ("Qty",        "數量",    60),
                ("Status",     "狀態",    65),
                ("StopLoss",   "停損",    70),
                ("TakeProfit", "停利",    70),
                ("RealizedPnL","損益",    75),
                ("OrderTime",  "掛單時間",90),
                ("AiReason",   "AI 理由", 140),
            };
            foreach (var (name, header, width) in cols)
            {
                dgOrders.Columns.Add(new DataGridViewTextBoxColumn
                { Name = name, HeaderText = header, FillWeight = width });
            }

            var pnlGrid = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(18, 20, 30) };
            var lblGridTitle = new Label
            {
                Text = "📋 模擬訂單記錄",
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(180, 200, 255),
                Dock = DockStyle.Top,
                Height = 28,
                Padding = new Padding(6, 4, 0, 0)
            };
            pnlGrid.Controls.Add(dgOrders);
            pnlGrid.Controls.Add(lblGridTitle);
            innerSplit.Panel1.Controls.Add(pnlGrid);

            // 執行日誌
            txtLog = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(10, 12, 20),
                ForeColor = Color.FromArgb(150, 220, 150),
                Font = new Font("Consolas", 9.5F),
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };
            var pnlLog = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(10, 12, 20) };
            var lblLogTitle = new Label
            {
                Text = "📟 執行日誌",
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(100, 220, 150),
                Dock = DockStyle.Top,
                Height = 28,
                Padding = new Padding(6, 4, 0, 0)
            };
            pnlLog.Controls.Add(txtLog);
            pnlLog.Controls.Add(lblLogTitle);
            innerSplit.Panel2.Controls.Add(pnlLog);
        }

        // ════════════════════════════════════════════════════════════════════════
        //  啟動 / 停止
        // ════════════════════════════════════════════════════════════════════════
        private async void BtnStart_Click(object sender, EventArgs e)
        {
            string ticker = txtTicker.Text.Trim().ToUpper();
            if (string.IsNullOrEmpty(ticker)) { Log("❌ 請輸入股票代號", Color.OrangeRed); return; }
            string key = txtApiKey.Text.Trim();
            if (string.IsNullOrEmpty(key)) { Log("❌ 請填入 API Key", Color.OrangeRed); return; }

            _config = new AutoTradeConfig
            {
                Ticker = ticker,
                DailyCapital = (double)numCapital.Value,
                MaxBuyOrders = (int)numMaxBuy.Value,
                MaxSellOrders = (int)numMaxSell.Value,
                MaxPositionPct = (double)numMaxPosPct.Value / 100.0,
                StopLossPct = (double)numStopLoss.Value / 100.0,
                TakeProfitPct = (double)numTakeProfit.Value / 100.0,
                PollIntervalSec = (int)numPollSec.Value,
                MinSignalScore = (double)numMinScore.Value,
                UseKellySize = chkKelly.Checked,
                AutoJournalLog = chkAutoJournal.Checked,
                TradingStyle = cbStyle.SelectedIndex == 0 ? "Day" : cbStyle.SelectedIndex == 2 ? "Position" : "Swing",
                TrailingStopPct = (double)numTrailingStop.Value / 100.0,
                MaxDailyLossPct = (double)numMaxDailyLoss.Value / 100.0,
            };
            _apiKey = key;
            _orderIdSeq = 1;

            _session = new AutoTradeSession
            {
                StartTime = DateTime.Now,
                IsRunning = true,
                InitialCapital = _config.DailyCapital,
                CurrentCapital = _config.DailyCapital
            };

            dgOrders.Rows.Clear();
            Log($"✅ 自動交易啟動  [{_config.Ticker}]  本金 ${_config.DailyCapital:N0}  輪詢 {_config.PollIntervalSec}s", Color.LightGreen);
            Log($"   停損 {_config.StopLossPct:P0}  停利 {_config.TakeProfitPct:P0}  AI 最低分 {_config.MinSignalScore:F0}  風格 {_config.TradingStyle}", Color.LightGray);

            btnStart.Enabled = false;
            btnStop.Enabled = true;
            lblStatus.Text = "▶ 執行中...";
            lblStatus.ForeColor = Color.LightGreen;

            _tradeCts = new CancellationTokenSource();
            try { await RunTradingLoop(_tradeCts.Token); }
            catch (OperationCanceledException) { Log("⏹ 自動交易已停止", Color.Orange); }
            catch (Exception ex) { Log($"❌ 例外：{ex.Message}", Color.OrangeRed); }
            finally
            {
                if (_session != null) { _session.IsRunning = false; _session.EndTime = DateTime.Now; }
                btnStart.Enabled = true;
                btnStop.Enabled = false;
                lblStatus.Text = "⏸ 已停止";
                lblStatus.ForeColor = Color.LightGray;
                Log($"📊 結算  損益 {_session?.TotalPnL:+N0;-N0} ({_session?.TotalPnLPct:+P2;-P2})  勝率 {_session?.WinRate:P1}", Color.Gold);
            }
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            _tradeCts?.Cancel();
        }

        // ════════════════════════════════════════════════════════════════════════
        //  主交易迴圈
        // ════════════════════════════════════════════════════════════════════════
        private async Task RunTradingLoop(CancellationToken ct)
        {
            // 追蹤今日買/賣單數量
            int todayBuyCount = 0, todaySellCount = 0;
            DateTime lastDate = DateTime.Today;

            // 目前持倉狀態
            double heldQty = 0;
            double holdEntry = 0;
            SimulatedOrder openBuy = null;  // 尚未成交的買單

            // 移動停損：追蹤持倉最高價
            double trailingHighPrice = 0;
            // 每日虧損追蹤
            double dailyStartCapital = _config.DailyCapital;
            bool dailyLossTriggered = false;

            while (!ct.IsCancellationRequested)
            {
                // 每日重置計數
                if (DateTime.Today != lastDate)
                {
                    todayBuyCount = 0; todaySellCount = 0; lastDate = DateTime.Today;
                    dailyStartCapital = _session.CurrentCapital;
                    dailyLossTriggered = false;
                    Log($"📅 新交易日 {DateTime.Today:MM/dd}，重置每日計數  本金基準 ${dailyStartCapital:N0}", Color.LightBlue);
                }

                // ── 每日最大虧損保護 ───────────────────────────────────────────
                if (!dailyLossTriggered && _config.MaxDailyLossPct > 0)
                {
                    double dailyLoss = (_session.CurrentCapital + _session.HoldingValue - dailyStartCapital) / dailyStartCapital;
                    if (dailyLoss <= -_config.MaxDailyLossPct)
                    {
                        dailyLossTriggered = true;
                        Log($"🚨 每日虧損上限觸發！今日虧損 {dailyLoss:P2} ≥ 設定上限 {_config.MaxDailyLossPct:P2}，停止今日交易", Color.OrangeRed);
                        ShowToast($"🚨 每日虧損上限 {_config.MaxDailyLossPct:P0} 觸發，今日停止交易", Color.OrangeRed);
                    }
                }

                lblStatus.Text = $"▶ 拉取數據中... {DateTime.Now:HH:mm:ss}";

                // ── Step 1：取得最新即時報價（用 Yahoo Finance）─────────────
                double currentPrice = 0;
                List<MarketData> recentData = null;
                try
                {
                    // 抓30日歷史 K 棒
                    var rawData = await YahooDataService.FetchHistoryAsync(_config.Ticker, "1mo", ct);
                    if (rawData?.Count > 5)
                    {
                        recentData = rawData;
                        IndicatorEngine.CalculateAll(recentData);
                        currentPrice = recentData.Last().Close;
                    }
                }
                catch (Exception ex) { Log($"⚠️ 數據拉取失敗：{ex.Message}", Color.Orange); }

                if (currentPrice <= 0 || recentData == null)
                {
                    await Task.Delay(TimeSpan.FromSeconds(_config.PollIntervalSec), ct);
                    continue;
                }

                Log($"📡 [{DateTime.Now:HH:mm:ss}] {_config.Ticker} = ${currentPrice:F2}", Color.FromArgb(150, 200, 255));

                // ── Step 1.5：VIX 動態風險調整（每 10 輪詢更新一次）────────────
                if (_mtfUpdateCounter % 10 == 0)
                {
                    try
                    {
                        var sentiment = await MarketSentimentService.FetchAsync();
                        if (sentiment != null && !sentiment.IsStale)
                        {
                            var riskLimit = SentimentAwareTradingEngine.CalcSentimentRiskLimit(sentiment);
                            double adjMaxPos = Math.Min(_config.MaxPositionPct, riskLimit.MaxPositionPct);
                            if (Math.Abs(adjMaxPos - _config.MaxPositionPct) > 0.01)
                            {
                                Log($"📊 VIX={sentiment.VIX:F1} [{riskLimit.Regime}] → 倉位上限調整 {_config.MaxPositionPct:P0} → {adjMaxPos:P0}", Color.Orange);
                                _config.MaxPositionPct = adjMaxPos;
                            }
                        }
                    }
                    catch { /* VIX 更新失敗不影響主流程 */ }
                }

                // ── Step 2：檢查現有持倉的停損/停利/移動停損 ─────────────────────
                if (heldQty > 0 && holdEntry > 0)
                {
                    // 更新移動停損高點
                    if (_config.TrailingStopPct > 0 && currentPrice > trailingHighPrice)
                        trailingHighPrice = currentPrice;

                    // 移動停損觸發：現價跌破最高點 × (1 - trailingStopPct)
                    bool hitTrailing = _config.TrailingStopPct > 0
                                       && trailingHighPrice > 0
                                       && currentPrice <= trailingHighPrice * (1 - _config.TrailingStopPct);

                    bool hitStop = currentPrice <= holdEntry * (1 - _config.StopLossPct);
                    bool hitTarget = currentPrice >= holdEntry * (1 + _config.TakeProfitPct);

                    if (hitTrailing && !hitStop)
                    {
                        // 移動停損觸發（優先）
                        double pnl = (currentPrice - holdEntry) * heldQty;
                        var tsOrder = CreateOrder("Sell", "Market", currentPrice, heldQty,
                            $"🔶 移動停損觸發 (最高點={trailingHighPrice:F2}→現價={currentPrice:F2})", holdEntry, 0, pnl);
                        FillOrder(tsOrder, currentPrice);
                        _session.Orders.Add(tsOrder);
                        AddOrderToGrid(tsOrder);
                        _session.CurrentCapital += currentPrice * heldQty;
                        _session.HoldingValue = 0;
                        if (pnl > 0) _session.WinCount++; else _session.LossCount++;
                        Log($"🔶 移動停損出場  最高={trailingHighPrice:F2}  現價={currentPrice:F2}  損益 {pnl:+N0;-N0}", Color.Orange);
                        if (_config.AutoJournalLog) WriteToJournal(tsOrder, holdEntry);
                        heldQty = 0; holdEntry = 0; openBuy = null; trailingHighPrice = 0;
                        todaySellCount++;
                    }
                    else if (hitStop || hitTarget)
                    {
                        string reason = hitStop ? "🔴 觸發停損" : "🟢 觸發停利";
                        double pnl = (currentPrice - holdEntry) * heldQty;

                        var sellOrder = CreateOrder("Sell", "Market", currentPrice, heldQty,
                            $"{reason} (進場={holdEntry:F2} → 現價={currentPrice:F2})", holdEntry, 0, pnl);

                        FillOrder(sellOrder, currentPrice);
                        _session.Orders.Add(sellOrder);
                        AddOrderToGrid(sellOrder);

                        _session.CurrentCapital += currentPrice * heldQty;
                        _session.HoldingValue = 0;
                        if (pnl > 0) _session.WinCount++; else _session.LossCount++;

                        Log($"{reason}  出場 {heldQty:N0} 股 @ {currentPrice:F2}  損益 {pnl:+N0;-N0}", hitStop ? Color.OrangeRed : Color.LightGreen);

                        if (_config.AutoJournalLog) WriteToJournal(sellOrder, holdEntry);

                        heldQty = 0; holdEntry = 0; openBuy = null; trailingHighPrice = 0;
                        todaySellCount++;
                    }
                }

                // ── Step 2.5：MTF 多時間框架信號更新（每 5 輪詢一次）─────────
                _mtfUpdateCounter++;
                if (_mtfUpdateCounter >= 5 || _mtfUpdateCounter == 1)
                {
                    _mtfUpdateCounter = 0;
                    lblStatus.Text = $"▶ 更新多時間框架信號... {DateTime.Now:HH:mm:ss}";
                    try
                    {
                        var mtf = await MultiTimeframeEngine.AnalyzeAsync(_config.Ticker, ct);
                        _mtfAlignmentScore = mtf.AlignmentScore;
                        Log($"🔀 MTF 共振分 {_mtfAlignmentScore:F0}/100  " +
                            $"週={mtf.Weekly_Trend} 日={mtf.Daily_Trend} 時={mtf.Hourly_Trend}", Color.FromArgb(150, 180, 255));
                    }
                    catch { /* MTF 更新失敗不影響主流程，保留上次值 */ }
                }

                // ── Step 3：AI 決策 ───────────────────────────────────────────
                lblStatus.Text = $"▶ AI 分析中... {DateTime.Now:HH:mm:ss}";

                var (aiAction, aiScore, aiReason) = await GetAiDecision(
                    recentData, currentPrice, heldQty, _apiKey, ct);

                // MTF 加權：最終信心分 = AI分 × 70% + MTF共振分 × 30%
                double effectiveScore = aiScore * 0.7 + _mtfAlignmentScore * 0.3;

                Log($"🤖 AI → {aiAction}  AI分 {aiScore:F0}  MTF {_mtfAlignmentScore:F0}  綜合 {effectiveScore:F0}/100  {aiReason.Substring(0, Math.Min(50, aiReason.Length))}", Color.FromArgb(220, 200, 100));

                // ── Step 4：條件判斷是否掛單（使用 MTF 加權後的 effectiveScore）──
                bool shouldBuy = aiAction == "Buy" && effectiveScore >= _config.MinSignalScore
                                  && heldQty == 0
                                  && todayBuyCount < _config.MaxBuyOrders
                                  && _session.CurrentCapital > 0
                                  && !dailyLossTriggered;

                bool shouldSell = aiAction == "Sell" && effectiveScore >= _config.MinSignalScore
                                  && heldQty > 0
                                  && todaySellCount < _config.MaxSellOrders;

                if (shouldBuy)
                {
                    // 計算買入數量（Kelly 或固定比例）
                    double posValue = _session.CurrentCapital * _config.MaxPositionPct;

                    // ── 台股必須以「張」為最小單位（1 張 = 1,000 股）──────────
                    bool isTwStock = _config.Ticker.EndsWith(".TW", StringComparison.OrdinalIgnoreCase)
                                  || _config.Ticker.EndsWith(".TWO", StringComparison.OrdinalIgnoreCase);
                    double qty;
                    if (isTwStock)
                    {
                        // 先算最多能買幾張，再乘以 1000 還原成股數
                        double lots = Math.Floor(posValue / (currentPrice * 1000.0));
                        if (lots < 1) lots = 1;            // 最少 1 張
                        qty = lots * 1000.0;               // 轉回股數（1,000 的整數倍）
                    }
                    else
                    {
                        qty = Math.Floor(posValue / currentPrice);
                        if (qty < 1) qty = 1;
                    }

                    double slPrice = currentPrice * (1 - _config.StopLossPct);
                    double tpPrice = currentPrice * (1 + _config.TakeProfitPct);

                    // 限價單：比現價低 0.1%（提高成交率）
                    double limitPx = Math.Round(currentPrice * 0.999, 2);

                    var order = CreateOrder("Buy", "Limit", limitPx, qty, aiReason, slPrice, tpPrice, 0);
                    _session.Orders.Add(order);
                    _session.TotalOrders++;
                    AddOrderToGrid(order);
                    string qtyLabel = isTwStock ? $"{qty / 1000:F0} 張 ({qty:N0} 股)" : $"{qty:N0} 股";
                    Log($"📥 掛買單 #{order.OrderId}  限價 ${limitPx:F2} × {qtyLabel}  停損 {slPrice:F2}  停利 {tpPrice:F2}", Color.FromArgb(100, 255, 150));

                    // ── 模擬撮合：限價單成交價 = 掛單價（limitPx）──────────
                    // 限價買單語意：「最多願意出 limitPx，成交就照 limitPx」
                    // 不應再從 currentPrice 加滑價，那會讓成交價偏離掛單價
                    await Task.Delay(500, ct);
                    double fillPx = limitPx;   // 成交價 = 掛單限價

                    FillOrder(order, fillPx);
                    _session.FilledOrders++;
                    UpdateOrderInGrid(order);

                    double cost = fillPx * qty;
                    _session.CurrentCapital -= cost;
                    _session.HoldingValue = fillPx * qty;
                    heldQty = qty; holdEntry = fillPx; openBuy = order;
                    trailingHighPrice = fillPx;  // 移動停損從進場價開始追蹤
                    todayBuyCount++;

                    Log($"✅ 買單成交 #{order.OrderId}  成交價 ${fillPx:F2}  持倉 {qtyLabel}  花費 ${cost:N0}", Color.LightGreen);
                    if (_config.AutoJournalLog) WriteToJournal(order, 0);
                }
                else if (shouldSell)
                {
                    double pnl = (currentPrice - holdEntry) * heldQty;
                    var order = CreateOrder("Sell", "Market", currentPrice, heldQty, aiReason, 0, 0, pnl);
                    _session.Orders.Add(order);
                    _session.TotalOrders++;
                    AddOrderToGrid(order);

                    double fillPx = Math.Round(currentPrice * (1 + (new Random().NextDouble() - 0.5) * 0.001), 2);
                    FillOrder(order, fillPx);
                    _session.FilledOrders++;
                    UpdateOrderInGrid(order);

                    _session.CurrentCapital += fillPx * heldQty;
                    _session.HoldingValue = 0;
                    if (pnl > 0) _session.WinCount++; else _session.LossCount++;
                    todaySellCount++;
                    heldQty = 0; holdEntry = 0;

                    Log($"✅ 賣單成交 #{order.OrderId}  成交價 ${fillPx:F2}  損益 {pnl:+N0;-N0}", pnl >= 0 ? Color.LightGreen : Color.OrangeRed);
                    if (_config.AutoJournalLog) WriteToJournal(order, openBuy?.FilledPrice ?? holdEntry);
                }
                else if (aiAction == "Hold" || aiScore < _config.MinSignalScore)
                {
                    string reason2 = aiScore < _config.MinSignalScore
                        ? $"信號分 {aiScore:F0} 低於閾值 {_config.MinSignalScore:F0}"
                        : "AI 建議持倉觀望";
                    Log($"⏸ {reason2}", Color.Gray);
                }

                // 更新持倉市值
                if (heldQty > 0)
                    _session.HoldingValue = currentPrice * heldQty;

                lblStatus.Text = $"▶ 等待下次輪詢... {DateTime.Now:HH:mm:ss}";
                await Task.Delay(TimeSpan.FromSeconds(_config.PollIntervalSec), ct);
            }
        }

        // ════════════════════════════════════════════════════════════════════════
        //  AI 決策（呼叫 LLM）
        // ════════════════════════════════════════════════════════════════════════
        private async Task<(string Action, double Score, string Reason)> GetAiDecision(
            List<MarketData> data, double currentPrice, double heldQty,
            string apiKey, CancellationToken ct)
        {
            try
            {
                var last = data.Last();
                string prompt = BuildAutoTradePrompt(data, currentPrice, heldQty, _config);
                string action = "Hold"; double score = 50; string reason = "預設觀望";

                var msgs = new List<object>
                {
                    new { role = "system", content =
                        "你是專業量化交易 AI，只輸出 JSON，格式：{\"action\":\"Buy|Sell|Hold\",\"score\":0-100,\"reason\":\"簡短理由\"}" },
                    new { role = "user", content = prompt }
                };

                string fullResp = "";
                await LLMService.StreamChat(apiKey, msgs, chunk => fullResp += chunk, ct);

                // 解析 JSON
                var m = System.Text.RegularExpressions.Regex.Match(
                    fullResp, @"\{[^}]+\}",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (m.Success)
                {
                    var doc = JsonDocument.Parse(m.Value).RootElement;
                    if (doc.TryGetProperty("action", out var a)) action = a.GetString() ?? "Hold";
                    if (doc.TryGetProperty("score", out var s)) score = s.GetDouble();
                    if (doc.TryGetProperty("reason", out var r)) reason = r.GetString() ?? "";
                }
                return (action, score, reason);
            }
            catch (Exception ex)
            {
                AppLogger.Log("AutoTrade AI 決策失敗", ex);
                return ("Hold", 0, $"AI 錯誤: {ex.Message}");
            }
        }

        private string BuildAutoTradePrompt(
            List<MarketData> data, double price, double heldQty, AutoTradeConfig cfg)
        {
            var last = data.Last();
            var prev5 = data.TakeLast(5).ToList();
            double atr = PositionSizingEngine.CalcATR(data, 14);
            double vol5 = prev5.Average(d => d.Volume);

            return $@"股票: {cfg.Ticker}  現價: ${price:F2}  風格: {cfg.TradingStyle}
目前持倉: {(heldQty > 0 ? $"{heldQty:N0} 股" : "空倉")}
技術指標: RSI={last.RSI:F1}  MACD柱={last.MACD_Hist:F3}    BB上軌={last.BB_Upper:F2}  BB下軌={last.BB_Lower:F2}
ATR={atr:F2}  5日均量={vol5:N0}  今日量={last.Volume:N0}  今日K={last.Pattern}
風控: 停損{cfg.StopLossPct:P0}  停利{cfg.TakeProfitPct:P0}  每日買單上限{cfg.MaxBuyOrders}次

請輸出 JSON，action 為 Buy/Sell/Hold，score 為 0-100 的信號強度（≥{cfg.MinSignalScore}才會執行），reason 為簡短中文理由（30字內）。
空倉時 Sell 無效。";
        }

        // ════════════════════════════════════════════════════════════════════════
        //  輔助方法
        // ════════════════════════════════════════════════════════════════════════
        private SimulatedOrder CreateOrder(
            string side, string type, double price, double qty, string reason,
            double sl, double tp, double pnl)
        {
            return new SimulatedOrder
            {
                OrderId = _orderIdSeq++,
                OrderTime = DateTime.Now,
                Ticker = _config.Ticker,
                Side = side,
                OrderType = type,
                LimitPrice = price,
                Qty = qty,
                Status = "Pending",
                AiReason = reason,
                StopLoss = sl,
                TakeProfit = tp,
                RealizedPnL = pnl
            };
        }

        private void FillOrder(SimulatedOrder order, double fillPx)
        {
            order.FilledPrice = fillPx;
            order.FilledTime = DateTime.Now;
            order.Status = "Filled";
            _session.Log.Add($"[{DateTime.Now:HH:mm:ss}] {order.Side} #{order.OrderId} 成交 @{fillPx:F2}");
        }

        private void WriteToJournal(SimulatedOrder order, double entryRef)
        {
            try
            {
                var all = ExcelJournalManager.LoadAll();
                int nextId = all.Count > 0 ? all.Max(x => x.Id) + 1 : 1;
                var entry = new TradeJournalEntry
                {
                    Id = nextId,
                    TradeDate = order.FilledTime ?? DateTime.Now,
                    Ticker = order.Ticker,
                    Direction = order.Side,
                    EntryPrice = order.Side == "Buy" ? order.FilledPrice : entryRef,
                    ExitPrice = order.Side == "Sell" ? order.FilledPrice : 0,
                    Quantity = order.Qty,
                    StopLossPrice = order.StopLoss,
                    TargetPrice = order.TakeProfit,
                    AiSuggestion = order.Side,
                    MyDecision = "AutoTrade",
                    Notes = $"[AI自動] {order.AiReason.Substring(0, Math.Min(100, order.AiReason.Length))}"
                };
                all.Add(entry);
                ExcelJournalManager.SaveAll(all);
            }
            catch (Exception ex) { AppLogger.Log("AutoTrade WriteToJournal 失敗", ex); }
        }

        private void AddOrderToGrid(SimulatedOrder o)
        {
            if (dgOrders.InvokeRequired) { dgOrders.Invoke(new Action(() => AddOrderToGrid(o))); return; }
            int idx = dgOrders.Rows.Add(
                o.OrderId, o.Side, o.OrderType,
                o.LimitPrice > 0 ? o.LimitPrice.ToString("F2") : "-",
                "-", o.Qty.ToString("N0"), o.Status,
                o.StopLoss > 0 ? o.StopLoss.ToString("F2") : "-",
                o.TakeProfit > 0 ? o.TakeProfit.ToString("F2") : "-",
                "-", o.OrderTime.ToString("HH:mm:ss"),
                o.AiReason.Length > 40 ? o.AiReason[..40] + "…" : o.AiReason);
            dgOrders.Rows[idx].DefaultCellStyle.ForeColor =
                o.Side == "Buy" ? Color.LightGreen : Color.LightCoral;
            dgOrders.FirstDisplayedScrollingRowIndex = idx;
        }

        private void UpdateOrderInGrid(SimulatedOrder o)
        {
            if (dgOrders.InvokeRequired) { dgOrders.Invoke(new Action(() => UpdateOrderInGrid(o))); return; }
            foreach (DataGridViewRow row in dgOrders.Rows)
            {
                if (row.Cells["OrderId"].Value?.ToString() == o.OrderId.ToString())
                {
                    row.Cells["FilledPrice"].Value = o.FilledPrice.ToString("F2");
                    row.Cells["Status"].Value = o.Status;
                    row.Cells["RealizedPnL"].Value = o.RealizedPnL != 0 ? o.RealizedPnL.ToString("+N0;-N0") : "-";
                    break;
                }
            }
        }

        private void RefreshStats()
        {
            if (_session == null) return;
            try
            {
                void Set(Label lbl, string text, Color color)
                {
                    if (lbl.InvokeRequired) lbl.Invoke(new Action(() => { lbl.Text = text; lbl.ForeColor = color; }));
                    else { lbl.Text = text; lbl.ForeColor = color; }
                }
                Set(lblCapital, $"💰 {_session.CurrentCapital:N0}", Color.White);
                Set(lblPnL, $"📈 {_session.TotalPnL:+N0;-N0} ({_session.TotalPnLPct:+P1;-P1})",
                    _session.TotalPnL >= 0 ? Color.LightGreen : Color.LightCoral);
                Set(lblWinRate, $"🎯 {_session.WinRate:P1} ({_session.WinCount}W/{_session.LossCount}L)", Color.Gold);
                Set(lblOrders, $"📋 {_session.FilledOrders}/{_session.TotalOrders}", Color.LightGray);
            }
            catch (Exception ex) { AppLogger.Log("UpdateStatusLabels 失敗", ex); }
        }

        private void Log(string msg, Color color)
        {
            if (txtLog.InvokeRequired) { txtLog.Invoke(new Action(() => Log(msg, color))); return; }
            txtLog.SelectionStart = txtLog.TextLength;
            txtLog.SelectionLength = 0;
            txtLog.SelectionColor = color;
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\n");
            txtLog.ScrollToCaret();
        }

        // ── UI 輔助工廠方法 ─────────────────────────────────────────────────
        private Label MakeLabel(string text) => new Label
        {
            Text = text,
            ForeColor = Color.LightGray,
            Font = new Font("Segoe UI", 9.5F),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft
        };

        private TextBox MakeTextBox(string text = "", bool password = false) => new TextBox
        {
            Text = text,
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(40, 42, 55),
            ForeColor = Color.White,
            Font = new Font("Consolas", 9.5F),
            BorderStyle = BorderStyle.FixedSingle,
            PasswordChar = password ? '•' : '\0'
        };

        private NumericUpDown MakeNumeric(decimal val, decimal min, decimal max, decimal inc,
            int decimals, bool thousands = false) => new NumericUpDown
            {
                Maximum = max, // 🔴 必須先設定 Maximum (放大天花板)
                Minimum = min, // 🔴 接著設定 Minimum (設定地板)
                Value = val,   // 🟢 最後才能設定 Value，這樣才不會超過範圍
                Increment = inc,
                DecimalPlaces = decimals,
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(40, 42, 55),
                ForeColor = Color.White,
                Font = new Font("Consolas", 9.5F),
                ThousandsSeparator = thousands
            };

        private Button MakeButton(string text, Color backColor)
        {
            var btn = new Button
            {
                Text = text,
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Height = 34,
                Width = 160,
                Margin = new Padding(0, 0, 8, 0)
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private Label MakeStat(string title, string value)
        {
            var lbl = new Label
            {
                Text = $"{title}\n{value}",
                Font = new Font("Segoe UI", 9.5F),
                ForeColor = Color.White,
                Width = 150,
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(28, 30, 44),
                Margin = new Padding(4),
                AutoSize = false,
                Height = 52
            };
            return lbl;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _tradeCts?.Cancel();
            _uiTimer?.Stop();
            base.OnFormClosed(e);
        }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  YahooDataService 擴充：FetchHistoryAsync（供 AutoTrade 用）
    //  ════════════════════════════════════════════════════════════════════════════
    // (此方法已在 Services.cs 的 YahooDataService 中，此處不重複)
}