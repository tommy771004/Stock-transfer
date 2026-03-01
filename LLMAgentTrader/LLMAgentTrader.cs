using System.Drawing.Drawing2D;
using System.Text;

namespace LLMAgentTrader
{
    public class QuantTraderForm : Form
    {
        // ── 全域 UI ──────────────────────────────────────────────────────────────
        private ComboBox cbMarketType;
        private ComboBox txtTicker;
        private TextBox txtApiKey;
        private NumericUpDown numDays;
        private Button btnRunAnalysis, btnOpenChat, btnNavToggle;
        private Panel pnlNav;           // 左側導覽面板
        private Panel pnlContent;       // 主內容區
        private Panel[] contentPanels;  // 7 個頁面 panel
        private Button[] navButtons;    // 導覽按鈕
        private int _currentPage = 0;
        private bool _navExpanded = false;
        private Panel pnlToast;         // Toast 通知浮層
        private Label lblToast;
        private System.Windows.Forms.Timer toastTimer;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel lblStatus;
        private ToolStripProgressBar progressBar;

        // 新功能 UI 欄位
        private TextBox txtRiskRewardResult;       // 風報比計算結果
        private Button btnCalcRR;                  // 計算停損停利
        private ComboBox cbRRDirection;            // 做多/做空選擇
        private DataGridView dgvAlerts;            // 警報追蹤表
        private Label lblAlertCount;               // 活躍警報數
        private System.Windows.Forms.Timer alertTimer; // 警報監控計時器
        // 市場情緒 (Tab 7)
        private Panel pnlFearGreed;
        private Label lblVIX, lblVIXChange, lblFGScore, lblFGLabel, lblSPChange, lblNdqChange, lblPosAdvice;
        private Button btnRefreshSentiment;
        private TextBox txtSentimentDetail;
        private MarketSentiment _lastSentiment = null;

        // 分頁1: 歷史策略
        private DataGridView dgvHistory;
        private QuantChartPanel historyChart;
        private TextBox txtDebateLog, txtNewsLog;
        private Label lblTotalReturn, lblWinRate, lblMDD, lblTradeCount;
        private Label lblSharpe, lblSortino, lblKelly;
        private Button btnMTF;
        private TextBox txtMTFResult;

        // 分頁2: 當沖
        private CheckBox chkLiveMode;
        private Label lblLivePrice, lblLiveChange, lblLiveHeartbeat, lblLiveSentiment;
        private QuantChartPanel liveChart;
        private OrderBookPanel orderBookPanel;
        private TextBox txtLiveAiDiagnosis;
        private Button btnLiveAiScan;
        private System.Windows.Forms.Timer liveTimer;
        private SemaphoreSlim netLock = new SemaphoreSlim(1, 1);

        // 分頁3: 投資組合
        private TextBox txtWatchlist;
        private Button btnRunPortfolio;
        private DataGridView dgvPortfolio;
        private TextBox txtPortfolioLog;

        // 分頁4: 提示詞
        private TextBox txtSystemPrompt, txtLivePrompt, txtPortfolioPrompt;
        private string histPromptFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "HistoricalPrompt.txt");
        private string livePromptFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "IntradayPrompt.txt");
        private string portfolioPromptFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PortfolioPrompt.txt");

        // 分頁5: 交易日誌
        private DataGridView dgvJournal;
        private Button btnJournalAdd, btnJournalEdit, btnJournalDelete,btnAutoAdd;
        private Label lblJournalStats;

        // 分頁6: 篩選器
        private TextBox txtScreenerTickers;
        private CheckBox chkRsiFilter, chkMacdPos, chkMacdCross, chkAboveEma50, chkAboveEma200, chkBBBreakout, chkBBOversold;
        private NumericUpDown numRsiMin, numRsiMax, numVolRatio;
        private Button btnRunScreener;
        private DataGridView dgvScreener;
        private TextBox txtScreenerLog;

        // 核心資料
        private List<MarketData> historyData = new List<MarketData>();
        private List<MarketData> currentLiveData = new List<MarketData>();
        private List<string> currentNews = new List<string>();
        private List<TradeJournalEntry> journalEntries = new List<TradeJournalEntry>();
        private MultiTimeframeSignal _lastMTFSignal = null;
        private string currentApiKey = "";
        private FloatingChatForm chatFormInstance = null;
        private string _lastTickTime = "", _currentCompanyName = "", _lastInfoTicker = "";
        private string _currentLiveTicker = "";   // ← 記錄當前當沖是哪支股票
        private double _lastTickPrice = -1, _lastAccumulatedVolume = -1, _currentPE = 0, _currentYield = 0;
        private System.Windows.Forms.Timer cacheTimer;
        // ── 取消分析用的 CancellationTokenSource
        private CancellationTokenSource _analysisCts = null;
        // ── 即時資料批次清除常數
        private const int LiveDataMaxSize = 5000;
        private const int LiveDataTrimBatch = 500;
        // ── ETF 相關
        private EtfInfo _etfInfo = null;
        private Panel pnlEtfCard = null;         // ETF 資訊卡（動態顯示）
        private Label lblEtfBadge = null;        // 頂部「ETF」標籤
        // ── 財報日提示
        private Label lblEarningsDate = null;
        // ── 板塊輪動面板（Tab 7）
        private Panel pnlSectorRotation = null;
        // ── 標的比較
        private string _compareTicker = "";
        private TextBox txtCompareTicker = null;
        private Button btnRunCompare = null;

        public QuantTraderForm()
        {
            InitializeUI();
            AutoScaleMode = AutoScaleMode.Dpi;
            liveTimer = new System.Windows.Forms.Timer { Interval = 5000 };
            liveTimer.Tick += async (s, e) => await RefreshLive();
            alertTimer = new System.Windows.Forms.Timer { Interval = 8000 };
            alertTimer.Tick += AlertTimer_Tick;
            // ← 修正：不在啟動時就 Start，改由 RefreshAlertGrid 偵測是否有活躍警報再決定啟停
            cacheTimer = new System.Windows.Forms.Timer { Interval = 5000 };
            cacheTimer.Tick += (s, e) => UpdateCachedTickers();
            cacheTimer.Start();
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            if (ClientRectangle.Width <= 0 || ClientRectangle.Height <= 0) return;
            using (var brush = new LinearGradientBrush(ClientRectangle, Color.FromArgb(10, 12, 18), Color.FromArgb(4, 5, 8), 60F))
                e.Graphics.FillRectangle(brush, ClientRectangle);
        }

        private void InitializeUI()
        {
            Text = "✨ Alpha-Twin v37.0 | Fib + MTF + Sharpe + Journal + Screener";
            Size = new Size(1620, 1020);
            MinimumSize = new Size(1400, 800);
            ForeColor = Color.White;
            Font = new Font("Segoe UI", 10.5F);

            // ── 頂部控制列 ────────────────────────────────────────────────────
            var pnlHeader = new Panel { Dock = DockStyle.Top, Height = 95, Padding = new Padding(25), BackColor = Color.FromArgb(30, 32, 40) };
            var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, BackColor = Color.Transparent, WrapContents = false };

            cbMarketType = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White, Font = new Font("Segoe UI", 11.5F, FontStyle.Bold), Width = 120, FlatStyle = FlatStyle.Flat };
            cbMarketType.Items.AddRange(new object[] { "🇹🇼 台股 (TWSE)", "🇺🇸 美股 (US)" });
            cbMarketType.SelectedIndex = 0;
            cbMarketType.SelectedIndexChanged += async (s, e) =>
            {
                if (cbMarketType.SelectedIndex == 0 && !txtTicker.Text.Contains(".")) txtTicker.Text = "2330.TW";
                if (cbMarketType.SelectedIndex == 1 && txtTicker.Text.Contains(".")) txtTicker.Text = "NVDA";
                await UpdateTickerInfoUI(); AutoLoadCachedData();
            };

            txtTicker = new ComboBox { Text = "2330.TW", BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White, Font = new Font("Segoe UI", 12F, FontStyle.Bold), Width = 120, FlatStyle = FlatStyle.Flat, DropDownStyle = ComboBoxStyle.DropDown };
            numDays = new NumericUpDown { Minimum = 30, Maximum = 1000, Value = 150, BorderStyle = BorderStyle.None, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White, Font = new Font("Segoe UI", 12F), Width = 80 };
            txtApiKey = new TextBox { PasswordChar = '•', BorderStyle = BorderStyle.None, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White, Font = new Font("Segoe UI", 12F), Width = 300 };

            btnRunAnalysis = MakeButton("✨ 執行分析", Color.FromArgb(0, 120, 212), BtnRunAnalysis_Click);
            btnOpenChat = MakeButton("💬 智能助理", Color.FromArgb(138, 43, 226), BtnOpenChat_Click);

            // 漢堡導覽切換鈕（最左側）
            btnNavToggle = new Button
            {
                Text = "☰",
                Width = 42,
                Height = 42,
                BackColor = Color.FromArgb(40, 42, 55),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 15F),
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 0, 10, 0)
            };
            btnNavToggle.FlatAppearance.BorderSize = 0;
            btnNavToggle.Click += (s, e) => ToggleNav();

            // ETF 快速選擇器
            var cbEtfQuick = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(45, 48, 55),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Width = 175,
                FlatStyle = FlatStyle.Flat
            };
            cbEtfQuick.Items.Add("⚡ 熱門 ETF");
            foreach (var (tk, lbl, _) in EtfService.WellKnownEtfs) cbEtfQuick.Items.Add($"{lbl} [{tk}]");
            cbEtfQuick.SelectedIndex = 0;
            cbEtfQuick.SelectedIndexChanged += (s, e) =>
            {
                if (cbEtfQuick.SelectedIndex <= 0) return;
                var (tk, _, _) = EtfService.WellKnownEtfs[cbEtfQuick.SelectedIndex - 1];
                txtTicker.Text = tk;
                cbMarketType.SelectedIndex = (tk.EndsWith(".TW") || tk.EndsWith(".TWO")) ? 0 : 1;
                DebounceHelper.Run("ticker", 50, async () =>
                {
                    if (!IsDisposed) { await UpdateTickerInfoUI(); AutoLoadCachedData(); }
                });
                cbEtfQuick.SelectedIndex = 0;   // 選後重置，保持提示文字
            };

            // 財報日提示標籤
            lblEarningsDate = new Label
            {
                Text = "",
                ForeColor = Color.Orange,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(4, 0, 0, 0),
                Visible = false
            };

            flow.Controls.AddRange(new Control[] {
                btnNavToggle,
                CreateComboWrapper(cbMarketType, "市場"),
                CreateInputWrapper(txtTicker, "標的代號"),
                CreateComboWrapper(cbEtfQuick, "ETF選擇"),
                CreateInputWrapper(numDays, "回測天數"),
                CreateInputWrapper(txtApiKey, "API Key"),
                btnRunAnalysis, btnOpenChat,
                lblEarningsDate
            });
            pnlHeader.Controls.Add(flow);
            Controls.Add(pnlHeader);

            // ── 主內容區 + 左側導覽 ──────────────────────────────────────────
            BuildNavAndContent();

            // 狀態列
            statusStrip = new StatusStrip { BackColor = Color.FromArgb(12, 14, 25), ForeColor = Color.Gray };
            lblStatus = new ToolStripStatusLabel { Text = "系統就緒" };
            progressBar = new ToolStripProgressBar { Visible = false, Width = 200 };
            statusStrip.Items.AddRange(new ToolStripItem[] { lblStatus, progressBar });
            Controls.Add(statusStrip);
            pnlHeader.SendToBack();

            // ── Toast 通知浮層 ───────────────────────────────────────────────
            pnlToast = new Panel
            {
                Size = new Size(380, 56),
                Visible = false,
                BackColor = Color.FromArgb(230, 30, 35, 45),
                Padding = new Padding(16, 10, 16, 10)
            };
            lblToast = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 10.5F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlToast.Controls.Add(lblToast);
            Controls.Add(pnlToast);
            pnlToast.BringToFront();
            toastTimer = new System.Windows.Forms.Timer { Interval = 4000 };
            toastTimer.Tick += (s, e) => { toastTimer.Stop(); pnlToast.Visible = false; };

            // ── 鍵盤快捷鍵 ──────────────────────────────────────────────────
            KeyPreview = true;
            KeyDown += (s, e) =>
            {
                if (e.Control)
                {
                    if (e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D7)
                    { ShowPage(e.KeyCode - Keys.D1); e.Handled = true; }
                    else if (e.KeyCode == Keys.Back)
                    { ToggleNav(); e.Handled = true; }
                }
                else if (e.KeyCode == Keys.Escape && _navExpanded)
                { ToggleNav(); e.Handled = true; }
            };

            // 事件
            txtApiKey.Leave += (s, e) => { string k = txtApiKey.Text.Trim(); if (!string.IsNullOrEmpty(k)) ApiKeyManager.Save(k); };
            // ← 修正：三個事件改用 Debounce 300ms，快速切換時只觸發最後一次，消除 race condition
            void OnTickerChanged()
            {
                DebounceHelper.Run("ticker", 300, async () =>
                {
                    if (IsDisposed) return;
                    string newTicker = txtTicker.Text.Trim().ToUpper();
                    if (newTicker != _currentLiveTicker) { _currentLiveTicker = ""; }
                    await UpdateTickerInfoUI();
                    AutoLoadCachedData();
                });
            }
            txtTicker.Leave += (s, e) => OnTickerChanged();
            txtTicker.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; ActiveControl = null; OnTickerChanged(); } };
            txtTicker.SelectedIndexChanged += (s, e) => { if (txtTicker.Focused) { ActiveControl = null; OnTickerChanged(); } };

            InitDataAsync();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (pnlToast != null)
                pnlToast.Location = new Point(ClientSize.Width - pnlToast.Width - 20,
                                              ClientSize.Height - pnlToast.Height - 30);
        }

        // ── Toast 通知 ───────────────────────────────────────────────────────
        private void ShowToast(string message, Color color = default)
        {
            if (InvokeRequired) { Invoke(new Action(() => ShowToast(message, color))); return; }
            lblToast.Text = message;
            pnlToast.BackColor = color == default ? Color.FromArgb(230, 30, 35, 45) : Color.FromArgb(230, color.R, color.G, color.B);
            pnlToast.Location = new Point(ClientSize.Width - pnlToast.Width - 20,
                                          ClientSize.Height - pnlToast.Height - 30);
            pnlToast.Visible = true;
            pnlToast.BringToFront();
            toastTimer.Stop(); toastTimer.Start();
        }

        // ── 左側導覽系統 ─────────────────────────────────────────────────────
        private void BuildNavAndContent()
        {
            var pnlMain = new Panel { Dock = DockStyle.Fill };
            Controls.Add(pnlMain);

            // 左側導覽面板（預設隱藏）
            pnlNav = new Panel
            {
                Dock = DockStyle.Left,
                Width = 220,
                Visible = false,
                BackColor = Color.FromArgb(14, 16, 26)
            };
            BuildNavPanel();

            // 主內容區
            pnlContent = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(10, 12, 18), Padding = new Padding(8, 6, 8, 8) };

            // 建立 7 個內容 Panel
            contentPanels = new Panel[7];
            for (int i = 0; i < 7; i++)
            {
                contentPanels[i] = new Panel { Dock = DockStyle.Fill, Visible = false, BackColor = Color.Transparent };
                pnlContent.Controls.Add(contentPanels[i]);
            }

            BuildTab1(contentPanels[0]);
            BuildTab2(contentPanels[1]);
            BuildTab3(contentPanels[2]);
            BuildTab4(contentPanels[3]);
            BuildTab5(contentPanels[4]);
            BuildTab6(contentPanels[5]);
            BuildTab7(contentPanels[6]);

            pnlMain.Controls.Add(pnlContent);
            pnlMain.Controls.Add(pnlNav);

            ShowPage(0);
        }

        private void BuildNavPanel()
        {
            // ── 頂部標題 ─────────────────────────────────────────────────────
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 56, BackColor = Color.FromArgb(10, 12, 20) };
            var lblTitle = new Label
            {
                Text = "⚡ ALPHA TWIN",
                Dock = DockStyle.Fill,
                ForeColor = Color.Gold,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };
            var btnClose = new Button
            {
                Text = "✕",
                Dock = DockStyle.Right,
                Width = 36,
                BackColor = Color.Transparent,
                ForeColor = Color.Gray,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 11F)
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) => ToggleNav();
            pnlTop.Controls.Add(lblTitle); pnlTop.Controls.Add(btnClose);

            // ── 分隔線 ────────────────────────────────────────────────────────
            var sep = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = Color.FromArgb(40, 44, 60) };

            // ── 導覽項目定義（圖示、名稱、頁面索引） ──────────────────────────
            var items = new (string icon, string name, int page)[]
            {
                ("📈", "歷史回測",   0),
                ("⚡", "當沖戰情室", 1),
                ("🎯", "投資組合",   2),
                ("🔍", "股票篩選器", 5),
                ("📓", "交易日誌",   4),
                ("📡", "市場情緒",   6),
                ("⚙️", "提示詞設定", 3),
            };

            navButtons = new Button[7];
            var pnlItems = new Panel { Dock = DockStyle.Top, Height = items.Length * 52 + 8, BackColor = Color.Transparent };
            for (int i = 0; i < items.Length; i++)
            {
                int pageIdx = items[i].page;
                int btnIdx = i;
                var btn = new Button
                {
                    Text = $"  {items[i].icon}  {items[i].name}",
                    TextAlign = ContentAlignment.MiddleLeft,
                    Location = new Point(10, 8 + i * 52),
                    Size = new Size(200, 44),
                    BackColor = Color.Transparent,
                    ForeColor = Color.FromArgb(180, 185, 200),
                    Font = new Font("Microsoft JhengHei UI", 10.5F),
                    FlatStyle = FlatStyle.Flat,
                    Tag = pageIdx,
                    Cursor = Cursors.Hand
                };
                btn.FlatAppearance.BorderSize = 0;
                btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(30, 80, 120, 200);
                int pi = pageIdx;
                btn.Click += (s, e) => { ShowPage(pi); ToggleNav(); };
                // 快捷鍵提示
                var sc = pi + 1;
                btn.Text = $"  {items[i].icon}  {items[i].name}          Ctrl+{sc}";
                btn.Font = new Font("Microsoft JhengHei UI", 10F);
                navButtons[i] = btn;
                pnlItems.Controls.Add(btn);
            }

            // ── 底部版本資訊 ──────────────────────────────────────────────────
            var lblVer = new Label
            {
                Text = "v37.0  |  Alpha-Twin",
                Dock = DockStyle.Bottom,
                Height = 32,
                ForeColor = Color.FromArgb(80, 90, 110),
                Font = new Font("Segoe UI", 8.5F),
                TextAlign = ContentAlignment.MiddleCenter
            };

            pnlNav.Controls.Add(lblVer);
            pnlNav.Controls.Add(pnlItems);
            pnlNav.Controls.Add(sep);
            pnlNav.Controls.Add(pnlTop);
        }

        private void ToggleNav()
        {
            _navExpanded = !_navExpanded;
            pnlNav.Visible = _navExpanded;
            btnNavToggle.BackColor = _navExpanded
                ? Color.FromArgb(0, 100, 180)
                : Color.FromArgb(40, 42, 55);
        }

        private void ShowPage(int pageIndex)
        {
            _currentPage = pageIndex;
            for (int i = 0; i < contentPanels.Length; i++)
                contentPanels[i].Visible = (i == pageIndex);

            // 更新導覽按鈕高亮（nav按鈕順序：0=頁0,1=頁1,2=頁2,3=頁5,4=頁4,5=頁6,6=頁3）
            int[] pageToNavBtn = { 0, 1, 2, 6, 4, 3, 5 }; // page→nav btn index
            if (navButtons != null)
            {
                for (int i = 0; i < navButtons.Length; i++)
                {
                    bool selected = (navButtons[i].Tag is int pi && pi == pageIndex);
                    navButtons[i].BackColor = selected
                        ? Color.FromArgb(0, 90, 170)
                        : Color.Transparent;
                    navButtons[i].ForeColor = selected
                        ? Color.White
                        : Color.FromArgb(180, 185, 200);
                    navButtons[i].Font = new Font("Microsoft JhengHei UI",
                        selected ? 10.5F : 10F,
                        selected ? FontStyle.Bold : FontStyle.Regular);
                }
            }

            // 觸發各頁面的進入邏輯
            if (pageIndex == 1) SwitchToLiveTab();
            if (pageIndex == 4) LoadJournal();
            if (pageIndex == 6) _ = RefreshSentimentAsync();

            // 更新狀態列頁面名稱
            string[] pageNames = { "歷史回測", "當沖戰情室", "投資組合", "提示詞設定", "交易日誌", "股票篩選器", "市場情緒" };
            //if (pageIndex < pageNames.Length)
            //    lblStatus.Text = $"📍 {pageNames[pageIndex]}   |   {lblStatus.Text.Split('|').LastOrDefault()?.Trim() ?? "就緒"}";
        }

        // ── 分頁1: 歷史策略回測 ──────────────────────────────────────────────
        private void BuildTab1(Panel t)
        {
            var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 950, BackColor = Color.Transparent, SplitterWidth = 8 };
            t.Controls.Add(split);

            // 左側：圖表 + 表格 + 統計
            var pnlL = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12), BackColor = Color.FromArgb(30, 32, 40) };
            historyChart = new QuantChartPanel { Dock = DockStyle.Top, Height = 340, BackColor = Color.Transparent };
            historyChart.OnDataPointClicked += (idx) => { if (idx >= 0 && idx < dgvHistory.Rows.Count) { dgvHistory.ClearSelection(); dgvHistory.Rows[idx].Selected = true; dgvHistory.FirstDisplayedScrollingRowIndex = Math.Max(0, idx - 5); } };

            // 多時間框架按鈕列
            var pnlMTF = new Panel { Dock = DockStyle.Top, Height = 42, BackColor = Color.FromArgb(22, 24, 32) };
            btnMTF = new Button { Text = "🔀 多時間框架分析 (週/日/時)", Dock = DockStyle.Left, Width = 260, Height = 38, BackColor = Color.FromArgb(80, 60, 150), ForeColor = Color.White, Font = new Font("Segoe UI", 10F, FontStyle.Bold), FlatStyle = FlatStyle.Flat, Margin = new Padding(4) };
            btnMTF.FlatAppearance.BorderSize = 0; btnMTF.Click += BtnMTF_Click;
            var chkFib = new CheckBox { Text = "📐 Fibonacci 回調線", Dock = DockStyle.Left, Width = 180, ForeColor = Color.Gold, Font = new Font("Segoe UI", 10F), CheckAlign = ContentAlignment.MiddleLeft, Checked = true, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(8, 0, 0, 0) };
            chkFib.CheckedChanged += (s, e) => { historyChart.ShowFibonacci = chkFib.Checked; historyChart.UpdateData(historyData, txtTicker.Text.Trim().ToUpper(), _currentCompanyName, _currentPE, _currentYield); };
            pnlMTF.Controls.Add(chkFib); pnlMTF.Controls.Add(btnMTF);

            dgvHistory = BuildDGV();
            dgvHistory.SelectionChanged += (s, e) => { if (dgvHistory.SelectedRows.Count > 0) historyChart.SetSelectedIndex(dgvHistory.SelectedRows[0].Index); };

            // 統計列
            var pnlStats = new Panel { Dock = DockStyle.Bottom, Height = 78, BackColor = Color.FromArgb(22, 25, 35) };
            lblTotalReturn = CreateStatLabel(pnlStats, "總報酬率", "0.0%", 10);
            lblWinRate = CreateStatLabel(pnlStats, "策略勝率", "0.0%", 135);
            lblMDD = CreateStatLabel(pnlStats, "最大回撤", "0.0%", 260);
            lblTradeCount = CreateStatLabel(pnlStats, "交易次數", "0", 385);
            lblSharpe = CreateStatLabel(pnlStats, "夏普比率", "-", 510);
            lblSortino = CreateStatLabel(pnlStats, "Sortino", "-", 635);
            lblKelly = CreateStatLabel(pnlStats, "Kelly 倉位", "-", 760);

            Panel pnlStatsRight = new Panel { Dock = DockStyle.Right, Width = 150, BackColor = Color.Transparent, Padding = new Padding(10, 25, 20, 25) };
            Button btnToggleChart = new Button { Text = "👁️ 隱藏線圖", Dock = DockStyle.Fill, BackColor = Color.FromArgb(60, 65, 75), ForeColor = Color.White, Font = new Font("Segoe UI", 10F, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnToggleChart.FlatAppearance.BorderSize = 0;
            btnToggleChart.Click += (s, e) => {
                historyChart.Visible = !historyChart.Visible;
                btnToggleChart.Text = historyChart.Visible ? "👁️ 隱藏線圖" : "👁️ 顯示線圖";
            };
            pnlStatsRight.Controls.Add(btnToggleChart);
            pnlStats.Controls.Add(pnlStatsRight);

            // 風報比計算列
            var pnlRR = new Panel { Dock = DockStyle.Bottom, Height = 88, BackColor = Color.FromArgb(20, 22, 30), Padding = new Padding(10, 6, 10, 6) };
            var rrTitle = new Label
            {
                Text = "📐 自動停損停利建議",
                Dock = DockStyle.Top,
                Height = 20,
                ForeColor = Color.Gold,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            var pnlRRCtrl = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 30, BackColor = Color.Transparent, WrapContents = false };
            cbRRDirection = new ComboBox
            {
                Width = 90,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(45, 48, 55),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            cbRRDirection.Items.AddRange(new object[] { "做多 Buy", "做空 Sell" }); cbRRDirection.SelectedIndex = 0;
            btnCalcRR = new Button
            {
                Text = "ATR法計算",
                Width = 110,
                Height = 30,
                BackColor = Color.FromArgb(0, 120, 100),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(6, 2, 4, 2),
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold)
            };
            btnCalcRR.FlatAppearance.BorderSize = 0;
            var btnCalcFib = new Button
            {
                Text = "Fib法計算",
                Width = 110,
                Height = 30,
                BackColor = Color.FromArgb(80, 60, 150),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(4, 2, 4, 2),
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold)
            };
            btnCalcFib.FlatAppearance.BorderSize = 0;
            btnCalcRR.Click += (s, e) => CalcRiskReward("ATR");
            btnCalcFib.Click += (s, e) => CalcRiskReward("Fib");
            pnlRRCtrl.Controls.AddRange(new Control[] { cbRRDirection, btnCalcRR, btnCalcFib });
            txtRiskRewardResult = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                BackColor = Color.FromArgb(15, 18, 25),
                ForeColor = Color.LightGoldenrodYellow,
                Font = new Font("Consolas", 10F),
                BorderStyle = BorderStyle.None
            };
            pnlRR.Controls.Add(txtRiskRewardResult); pnlRR.Controls.Add(pnlRRCtrl); pnlRR.Controls.Add(rrTitle);

            pnlL.Controls.Add(dgvHistory); pnlL.Controls.Add(pnlRR); pnlL.Controls.Add(pnlStats); pnlL.Controls.Add(pnlMTF); pnlL.Controls.Add(historyChart);
            split.Panel1.Controls.Add(pnlL);

            // 右側：ETF資訊卡（條件顯示）+ AI辯論 + MTF結果 + 新聞 + 比較工具
            var pnlR = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, ColumnCount = 1, BackColor = Color.FromArgb(28, 30, 38) };
            pnlR.RowStyles.Add(new RowStyle(SizeType.AutoSize));         // ETF Card（折疊）
            pnlR.RowStyles.Add(new RowStyle(SizeType.Percent, 55F));
            pnlR.RowStyles.Add(new RowStyle(SizeType.Percent, 22F));
            pnlR.RowStyles.Add(new RowStyle(SizeType.Percent, 23F));

            // ── ETF 資訊卡 ────────────────────────────────────────────────
            pnlEtfCard = BuildEtfCardPanel();
            pnlR.Controls.Add(pnlEtfCard, 0, 0);

            txtDebateLog = MakeTextBox(Color.FromArgb(180, 255, 200), new Font("Consolas", 11F));
            txtMTFResult = MakeTextBox(Color.Gold, new Font("Consolas", 10.5F));
            txtNewsLog = MakeTextBox(Color.LightGray, new Font("Microsoft JhengHei UI", 10.5F));

            pnlR.Controls.Add(MakeLabeledPanel("🤖 AI 歷史辯論與決策", txtDebateLog, Color.White), 0, 1);
            pnlR.Controls.Add(MakeLabeledPanel("🔀 多時間框架共振分析", txtMTFResult, Color.Gold), 0, 2);
            pnlR.Controls.Add(MakeLabeledPanel("📰 重大市場情緒與新聞", txtNewsLog, Color.LightSkyBlue), 0, 3);
            split.Panel2.Controls.Add(pnlR);

            // ── 比較工具（底部小列）────────────────────────────────────────
            var pnlCompare = new Panel { Dock = DockStyle.Bottom, Height = 38, BackColor = Color.FromArgb(18, 20, 30), Padding = new Padding(8, 4, 8, 4) };
            var compareFlow = new FlowLayoutPanel { Dock = DockStyle.Fill, BackColor = Color.Transparent, WrapContents = false };
            var lblComp = new Label { Text = "📊 比較:", ForeColor = Color.LightGray, Font = new Font("Segoe UI", 9.5F), AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Margin = new Padding(0, 4, 4, 0) };
            txtCompareTicker = new TextBox { Width = 80, Height = 26, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White, Font = new Font("Segoe UI", 10F), BorderStyle = BorderStyle.FixedSingle };
            var cbCompRange = new ComboBox { Width = 70, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F) };
            cbCompRange.Items.AddRange(new object[] { "3mo", "6mo", "1y", "2y" }); cbCompRange.SelectedIndex = 2;
            btnRunCompare = new Button { Text = "▶ 比較走勢", Width = 100, Height = 26, BackColor = Color.FromArgb(100, 60, 160), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold) };
            btnRunCompare.FlatAppearance.BorderSize = 0;
            btnRunCompare.Click += async (s, e) => await RunComparisonAsync(
                txtTicker.Text.Trim().ToUpper(),
                txtCompareTicker.Text.Trim().ToUpper(),
                cbCompRange.Text);
            compareFlow.Controls.AddRange(new Control[] { lblComp, txtCompareTicker, cbCompRange, btnRunCompare });
            pnlCompare.Controls.Add(compareFlow);
            split.Panel1.Controls.Add(pnlCompare);
        }

        // ── 分頁2: 當沖戰情室 ──────────────────────────────────────────────
        private void BuildTab2(Panel t)
        {
            var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 950, BackColor = Color.Transparent, SplitterWidth = 8 };
            t.Controls.Add(split);

            var pnlOB = new Panel { Dock = DockStyle.Right, Width = 260, BackColor = Color.FromArgb(30, 32, 40) };
            orderBookPanel = new OrderBookPanel { Dock = DockStyle.Fill };
            pnlOB.Controls.Add(orderBookPanel);

            var pnlLiveChart = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12), BackColor = Color.FromArgb(30, 32, 40) };
            var pnlLH = new Panel { Dock = DockStyle.Top, Height = 82, BackColor = Color.Transparent };

            lblLivePrice = new Label 
            { 
                Text = "$ 0.00",
                Dock = DockStyle.Left,
                Width = 230,
                Font = new Font("Consolas", 36F, FontStyle.Bold),
                TextAlign = ContentAlignment.BottomLeft,
                ForeColor = Color.White,          // 明確顏色
                BackColor = Color.Transparent,
                Margin = new Padding(8, 0, 8, 0)
            };
            lblLiveChange = new Label { Text = "+0.00%", Dock = DockStyle.Left, Width = 220, Font = new Font("Segoe UI", 16F), TextAlign = ContentAlignment.BottomLeft, Padding = new Padding(0, 0, 0, 8) };
            lblLiveSentiment = new Label { Text = "⚪ 累積中", Dock = DockStyle.Left, Width = 260, Font = new Font("Segoe UI", 13F, FontStyle.Bold), ForeColor = Color.Gray, TextAlign = ContentAlignment.BottomLeft, Padding = new Padding(0, 0, 0, 8) };

            var pnlCtrl = new Panel { Dock = DockStyle.Right, Width = 280, BackColor = Color.Transparent };
            lblLiveHeartbeat = new Label { Text = "等待...", Dock = DockStyle.Top, Height = 28, ForeColor = Color.Gray, Font = new Font("Consolas", 10F), TextAlign = ContentAlignment.MiddleRight };
            chkLiveMode = new CheckBox { Text = "🛡️ 自動刷新 5s", Dock = DockStyle.Top, Height = 38, ForeColor = Color.White, Font = new Font("Segoe UI", 12F), CheckAlign = ContentAlignment.MiddleRight, TextAlign = ContentAlignment.MiddleRight };
            chkLiveMode.CheckedChanged += async (s, e) =>
            {
                liveTimer.Enabled = chkLiveMode.Checked;
                if (chkLiveMode.Checked)
                {
                    lblLiveHeartbeat.Text = "刷新中...";
                    lblLiveHeartbeat.ForeColor = Color.Gold;
                    string tk = txtTicker.Text.Trim().ToUpper();

                    // 股票與當前不同時，觸發切換清空
                    if (tk != _currentLiveTicker) SwitchToLiveTab();

                    // 若還沒有今日快取才載入
                    if (currentLiveData.Count == 0)
                    {
                        var cached = ExcelIntradayCacheManager.LoadTodayData(tk);
                        if (cached.Count > 0)
                        {
                            currentLiveData = cached;
                            _lastTickTime = currentLiveData.Last().Date.ToString("yyyy-MM-dd HH:mm:ss");
                            _lastTickPrice = currentLiveData.Last().Close;
                            _lastAccumulatedVolume = currentLiveData.Last().Volume;
                        }
                    }
                    await RefreshLive();
                }
                else
                {
                    lblLiveHeartbeat.Text = "已暫停";
                    lblLiveHeartbeat.ForeColor = Color.Gray;
                }
            };
            pnlCtrl.Controls.Add(chkLiveMode); pnlCtrl.Controls.Add(lblLiveHeartbeat);
            pnlLH.Controls.Add(lblLiveSentiment); pnlLH.Controls.Add(lblLiveChange); pnlLH.Controls.Add(lblLivePrice); pnlLH.Controls.Add(pnlCtrl);

            liveChart = new QuantChartPanel { Dock = DockStyle.Fill, IsIntraday = true, ShowFibonacci = false };
            pnlLiveChart.Controls.Add(liveChart); pnlLiveChart.Controls.Add(pnlLH);

            var pnlCore = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
            pnlCore.Controls.Add(pnlLiveChart); pnlCore.Controls.Add(pnlOB);
            split.Panel1.Controls.Add(pnlCore);

            var pnlAI = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15), BackColor = Color.FromArgb(28, 30, 38) };
            btnLiveAiScan = new Button { Text = "⚡ AI 診斷即時動能 + VWAP + Fib", Dock = DockStyle.Top, Height = 48, BackColor = Color.FromArgb(200, 80, 0), ForeColor = Color.White, Font = new Font("Segoe UI", 12F, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btnLiveAiScan.FlatAppearance.BorderSize = 0; btnLiveAiScan.Click += BtnLiveAiScan_Click;
            txtLiveAiDiagnosis = MakeTextBox(Color.LightGoldenrodYellow, new Font("Microsoft JhengHei UI", 11.5F));
            pnlAI.Controls.Add(txtLiveAiDiagnosis);
            pnlAI.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 12 });
            pnlAI.Controls.Add(btnLiveAiScan);
            split.Panel2.Controls.Add(pnlAI);
        }

        // ── 分頁3: 投資組合 ──────────────────────────────────────────────────
        private void BuildTab3(Panel t)
        {
            var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 680, BackColor = Color.Transparent, SplitterWidth = 8 };
            var pnlL = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12), BackColor = Color.FromArgb(30, 32, 40) };
            var pnlInput = new Panel { Dock = DockStyle.Top, Height = 72, BackColor = Color.Transparent };
            txtWatchlist = new TextBox { Text = "2330.TW, 2317.TW, NVDA, TSLA", BorderStyle = BorderStyle.None, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White, Font = new Font("Segoe UI", 12F, FontStyle.Bold), Width = 430 };
            btnRunPortfolio = MakeButton("🚀 雷達掃描", Color.FromArgb(0, 120, 212), BtnRunPortfolio_Click);
            var ww = CreateInputWrapper(txtWatchlist, "多檔清單 (逗號分隔)"); ww.Dock = DockStyle.Left;
            var pnlBtn = new Panel { Dock = DockStyle.Left, Width = 200, Padding = new Padding(12, 16, 0, 0) };
            pnlBtn.Controls.Add(btnRunPortfolio);
            pnlInput.Controls.Add(pnlBtn); pnlInput.Controls.Add(ww);
            dgvPortfolio = BuildDGV();
            pnlL.Controls.Add(dgvPortfolio); pnlL.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 12 }); pnlL.Controls.Add(pnlInput);
            split.Panel1.Controls.Add(pnlL);
            var pnlR = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15), BackColor = Color.FromArgb(28, 30, 38) };
            txtPortfolioLog = MakeTextBox(Color.Plum, new Font("Microsoft JhengHei UI", 11.5F));
            pnlR.Controls.Add(txtPortfolioLog);
            pnlR.Controls.Add(new Label { Text = "🧠 AI MPT 配置建議", Dock = DockStyle.Top, Height = 38, ForeColor = Color.White, Font = new Font("Segoe UI", 12F, FontStyle.Bold) });
            split.Panel2.Controls.Add(pnlR);
            t.Controls.Add(split);
        }

        // ── 分頁4: 提示詞設定 ────────────────────────────────────────────────
        private void BuildTab4(Panel t)
        {
            var tlp = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, BackColor = Color.Transparent };
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 33.3F));
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 33.3F));
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 33.3F));
            txtSystemPrompt = new TextBox { Multiline = true, Dock = DockStyle.Fill, BackColor = Color.FromArgb(18, 20, 28), ForeColor = Color.LightGoldenrodYellow, Font = new Font("Consolas", 12F), BorderStyle = BorderStyle.None };
            txtLivePrompt = new TextBox { Multiline = true, Dock = DockStyle.Fill, BackColor = Color.FromArgb(18, 20, 28), ForeColor = Color.LightSkyBlue, Font = new Font("Consolas", 12F), BorderStyle = BorderStyle.None };
            txtPortfolioPrompt = new TextBox { Multiline = true, Dock = DockStyle.Fill, BackColor = Color.FromArgb(18, 20, 28), ForeColor = Color.Plum, Font = new Font("Consolas", 12F), BorderStyle = BorderStyle.None };
            tlp.Controls.Add(MakeLabeledPanel("歷史回測 AI Prompt:", txtSystemPrompt, Color.LightGoldenrodYellow, Color.FromArgb(30, 32, 40)), 0, 0);
            tlp.Controls.Add(MakeLabeledPanel("即時當沖 AI Prompt:", txtLivePrompt, Color.LightSkyBlue, Color.FromArgb(30, 32, 40)), 0, 1);
            tlp.Controls.Add(MakeLabeledPanel("投資組合 AI Prompt:", txtPortfolioPrompt, Color.Plum, Color.FromArgb(30, 32, 40)), 0, 2);

            // ── 底部工具列（儲存 + 模型選擇）──────────────────────────────
            var pnlSave = new Panel { Dock = DockStyle.Bottom, Height = 65, Padding = new Padding(14, 12, 14, 12), BackColor = Color.FromArgb(20, 22, 30) };

            // 模型選擇下拉（修正：原本 model 硬編碼，現在可在 UI 切換）
            var lblModel = new Label
            {
                Text = "🤖 AI 模型:",
                ForeColor = Color.LightGray,
                TextAlign = ContentAlignment.MiddleLeft,
                Width = 80,
                Dock = DockStyle.Left,
                Font = new Font("Segoe UI", 10F)
            };
            var cbModel = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(40, 42, 55),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10.5F, FontStyle.Bold),
                Width = 250,
                Dock = DockStyle.Left,
                FlatStyle = FlatStyle.Flat
            };
            foreach (var (id, label) in LlmConfig.AvailableModels)
                cbModel.Items.Add(label);
            // 預設選中目前設定的模型
            int modelIdx = Array.FindIndex(LlmConfig.AvailableModels, m => m.Id == LlmConfig.CurrentModel);
            cbModel.SelectedIndex = modelIdx >= 0 ? modelIdx : 0;
            cbModel.SelectedIndexChanged += (s, e) =>
            {
                if (cbModel.SelectedIndex >= 0 && cbModel.SelectedIndex < LlmConfig.AvailableModels.Length)
                {
                    LlmConfig.CurrentModel = LlmConfig.AvailableModels[cbModel.SelectedIndex].Id;
                    ShowToast($"✅ 模型已切換為 {LlmConfig.AvailableModels[cbModel.SelectedIndex].Label}",
                              Color.FromArgb(20, 80, 120));
                }
            };

            var btnSave = MakeButton("💾 儲存提示詞", Color.FromArgb(0, 120, 212), BtnSavePrompts_Click);
            btnSave.Dock = DockStyle.Right; btnSave.Width = 180;

            var flowModel = new FlowLayoutPanel { Dock = DockStyle.Left, BackColor = Color.Transparent, AutoSize = true, Padding = new Padding(0, 2, 0, 0) };
            flowModel.Controls.Add(lblModel);
            flowModel.Controls.Add(cbModel);
            pnlSave.Controls.Add(btnSave);
            pnlSave.Controls.Add(flowModel);

            // ── Alpha Vantage Key 設定列 ──────────────────────────────────
            var pnlAV = new Panel { Dock = DockStyle.Bottom, Height = 38, Padding = new Padding(14, 6, 14, 6), BackColor = Color.FromArgb(16, 18, 26) };
            var lblAV = new Label
            {
                Text = "🔑 Alpha Vantage Key (免費 PE/殖利率備援):",
                ForeColor = Color.LightGray,
                Width = 280,
                Font = new Font("Segoe UI", 9.5F),
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Left
            };
            var txtAVKey = new TextBox
            {
                Width = 230,
                Height = 24,
                Dock = DockStyle.Left,
                BackColor = Color.FromArgb(40, 42, 55),
                ForeColor = Color.LightYellow,
                Font = new Font("Consolas", 10F),
                BorderStyle = BorderStyle.FixedSingle,
                PasswordChar = '•',
                Text = AlphaVantageKeyManager.Key
            };
            var btnAVSave = new Button
            {
                Text = "儲存",
                Width = 60,
                Height = 24,
                Dock = DockStyle.Left,
                BackColor = Color.FromArgb(0, 100, 100),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Margin = new Padding(4, 0, 0, 0)
            };
            btnAVSave.FlatAppearance.BorderSize = 0;
            btnAVSave.Click += (s, e) =>
            {
                AlphaVantageKeyManager.Save(txtAVKey.Text.Trim());
                ShowToast("✅ Alpha Vantage Key 已儲存", Color.FromArgb(0, 80, 40));
            };
            var lnkAV = new LinkLabel
            {
                Text = "免費申請",
                Dock = DockStyle.Left,
                LinkColor = Color.SkyBlue,
                Font = new Font("Segoe UI", 9F),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(6, 0, 0, 0),
                Width = 60
            };
            lnkAV.Click += (s, e) =>
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                { FileName = "https://www.alphavantage.co/support/#api-key", UseShellExecute = true });
            var avFlow = new FlowLayoutPanel { Dock = DockStyle.Fill, BackColor = Color.Transparent, WrapContents = false };
            avFlow.Controls.AddRange(new Control[] { lblAV, txtAVKey, btnAVSave, lnkAV });
            pnlAV.Controls.Add(avFlow);

            t.Controls.Add(tlp); t.Controls.Add(pnlAV); t.Controls.Add(pnlSave);
            LoadPrompts();
        }

        // ── 分頁5: 交易日誌 ──────────────────────────────────────────────────
        private void BuildTab5(Panel t)
        {
            var split5 = new SplitContainer
            {
                Dock = DockStyle.Fill,
                SplitterDistance = 700,
                BackColor = Color.Transparent,
                SplitterWidth = 8
            };
            t.Controls.Add(split5);

            // 左側：日誌表格 + 工具列
            var pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10), BackColor = Color.FromArgb(28, 30, 38) };
            var pnlTools = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(22, 24, 32) };
            btnJournalAdd = MakeButton("➕ 新增", Color.FromArgb(0, 150, 80), BtnJournalAdd_Click); btnJournalAdd.Margin = new Padding(5); btnJournalAdd.Width = 120;
            btnAutoAdd = MakeButton("模擬自動交易設定", Color.FromArgb(0, 150, 80), btnOpenAIAutoTrade_Click); btnAutoAdd.Margin = new Padding(5); btnAutoAdd.Width = 140;
            btnJournalEdit = MakeButton("✏️ 修改/平倉", Color.FromArgb(100, 100, 0), BtnJournalEdit_Click); btnJournalEdit.Margin = new Padding(5); btnJournalEdit.Width = 160;
            btnJournalDelete = MakeButton("🗑️ 刪除", Color.FromArgb(160, 0, 0), BtnJournalDelete_Click); btnJournalDelete.Margin = new Padding(5); btnJournalDelete.Width = 100;
            lblJournalStats = new Label
            {
                Dock = DockStyle.Right,
                Width = 200,
                ForeColor = Color.LightGoldenrodYellow,
                Font = new Font("Consolas", 10F),
                TextAlign = ContentAlignment.MiddleRight,
                Text = "統計: 載入中..."
            };
            var fJournal = new FlowLayoutPanel { Dock = DockStyle.Left, BackColor = Color.Transparent, AutoSize = true, WrapContents = false };
            fJournal.Controls.AddRange(new Control[] { btnJournalAdd, btnAutoAdd, btnJournalEdit, btnJournalDelete });
            pnlTools.Controls.Add(lblJournalStats); pnlTools.Controls.Add(fJournal);
            dgvJournal = BuildDGV();
            pnlLeft.Controls.Add(dgvJournal); pnlLeft.Controls.Add(pnlTools);
            split5.Panel1.Controls.Add(pnlLeft);

            // 右側：止盈止損警報追蹤
            var pnlAlert = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12), BackColor = Color.FromArgb(22, 25, 35) };
            var pnlAlertHdr = new Panel { Dock = DockStyle.Top, Height = 45, BackColor = Color.Transparent };
            var lblAlertTitle = new Label
            {
                Text = "🚨 止盈止損警報追蹤",
                Dock = DockStyle.Left,
                Width = 250,
                ForeColor = Color.OrangeRed,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold)
            };
            lblAlertCount = new Label
            {
                Text = "活躍: 0 筆",
                Dock = DockStyle.Right,
                Width = 120,
                ForeColor = Color.Gold,
                Font = new Font("Consolas", 11F),
                TextAlign = ContentAlignment.MiddleRight
            };
            var btnDismiss = new Button
            {
                Text = "✅ 清除已觸發",
                Dock = DockStyle.Right,
                Width = 130,
                Height = 36,
                BackColor = Color.FromArgb(80, 80, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 4, 8, 4)
            };
            btnDismiss.FlatAppearance.BorderSize = 0;
            btnDismiss.Click += (s, e) => {
                foreach (DataGridViewRow row in dgvAlerts.SelectedRows)
                {
                    if (row.Cells["JournalId"].Value is int jid) AlertEngine.DismissAlert(jid);
                }
                RefreshAlertGrid();
            };
            pnlAlertHdr.Controls.Add(btnDismiss); pnlAlertHdr.Controls.Add(lblAlertCount); pnlAlertHdr.Controls.Add(lblAlertTitle);

            var lblAlertHint = new Label
            {
                Text = "📌 新增交易並填入「停損價」和「目標價」後，系統每 8 秒自動比對即時報價觸發警報。",
                Dock = DockStyle.Top,
                Height = 32,
                ForeColor = Color.LightGray,
                Font = new Font("Segoe UI", 9F)
            };

            dgvAlerts = BuildDGV();
            pnlAlert.Controls.Add(dgvAlerts); pnlAlert.Controls.Add(lblAlertHint); pnlAlert.Controls.Add(pnlAlertHdr);
            split5.Panel2.Controls.Add(pnlAlert);

            // ── 底部: 績效曲線圖 ─────────────────────────────────────────────
            var pnlEquity = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 130,
                BackColor = Color.FromArgb(16, 18, 26),
                Padding = new Padding(10, 6, 10, 6)
            };
            var lblEqTitle = new Label
            {
                Text = "📉 累積損益曲線 (交易日誌)",
                Dock = DockStyle.Top,
                Height = 22,
                ForeColor = Color.Gold,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold)
            };
            var equityChart = new EquityCurvePanel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
            equityChart.Name = "equityChart";
            pnlEquity.Controls.Add(equityChart);
            pnlEquity.Controls.Add(lblEqTitle);
            t.Controls.Add(pnlEquity);
            t.Controls.Add(split5); // split5 先 fill，equityChart bottom
        }
        private void btnOpenAIAutoTrade_Click(object sender, EventArgs e)
        {
            AutoTradeForm aiForm = new AutoTradeForm();
            aiForm.Show(); // 開啟非獨佔視窗，讓你同時可以看到其他主程式數據
        }
        // ── 分頁7: 市場情緒儀表板 ───────────────────────────────────────────────
        private void BuildTab7(Panel t)
        {
            var split7 = new SplitContainer
            {
                Dock = DockStyle.Fill,
                SplitterDistance = 420,
                BackColor = Color.Transparent,
                SplitterWidth = 8,
                Orientation = Orientation.Vertical
            };
            t.Controls.Add(split7);

            // 左側：儀表板指標
            var pnlDash = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 15, 20, 15),
                BackColor = Color.FromArgb(18, 20, 28),
                AutoScroll = true
            };

            var lblTitle = new Label
            {
                Text = "📡 市場整體情緒儀表板",
                Dock = DockStyle.Top,
                Height = 40,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 14F, FontStyle.Bold)
            };

            btnRefreshSentiment = new Button
            {
                Text = "🔄 立即刷新",
                Dock = DockStyle.Top,
                Height = 42,
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat
            };
            btnRefreshSentiment.FlatAppearance.BorderSize = 0;
            btnRefreshSentiment.Click += async (s, e) => await RefreshSentimentAsync(force: true);

            // 各指標 Label
            lblVIX = MakeSentimentLabel("VIX 恐慌指數", "-", Color.OrangeRed);
            lblVIXChange = MakeSentimentLabel("VIX 變化", "-", Color.LightSalmon);
            lblFGScore = MakeSentimentLabel("Fear & Greed", "- / 100", Color.Gold);
            lblFGLabel = MakeSentimentLabel("市場情緒", "-", Color.Gold);
            lblSPChange = MakeSentimentLabel("S&P 500", "-", Color.LightGreen);
            lblNdqChange = MakeSentimentLabel("Nasdaq", "-", Color.LightSkyBlue);
            lblPosAdvice = new Label
            {
                Dock = DockStyle.Top,
                Height = 60,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold),
                Text = "倉位建議: 載入中...",
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(0, 8, 0, 0)
            };

            pnlFearGreed = new Panel { Dock = DockStyle.Top, Height = 28, BackColor = Color.Transparent };
            pnlFearGreed.Paint += (s, e) => DrawFearGreedBar(e.Graphics, pnlFearGreed.Width, pnlFearGreed.Height);

            pnlDash.Controls.Add(lblPosAdvice);
            pnlDash.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 8 });
            pnlDash.Controls.Add(pnlFearGreed);
            pnlDash.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 8 });
            pnlDash.Controls.Add(lblFGLabel); pnlDash.Controls.Add(lblFGScore);
            pnlDash.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 8 });
            pnlDash.Controls.Add(lblNdqChange); pnlDash.Controls.Add(lblSPChange);
            pnlDash.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 8 });
            pnlDash.Controls.Add(lblVIXChange); pnlDash.Controls.Add(lblVIX);
            pnlDash.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 12 });
            pnlDash.Controls.Add(btnRefreshSentiment);
            pnlDash.Controls.Add(lblTitle);
            split7.Panel1.Controls.Add(pnlDash);

            // 右側：板塊輪動（下） + 詳細說明（上）
            var pnlDetail = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15), BackColor = Color.FromArgb(22, 25, 35) };
            // 板塊輪動面板（右側下方）
            var sectorPanel = BuildSectorRotationPanel();
            pnlDetail.Controls.Add(sectorPanel);
            txtSentimentDetail = MakeTextBox(Color.LightYellow, new Font("Microsoft JhengHei UI", 11.5F));
            txtSentimentDetail.Text = "🔄 點擊「立即刷新」取得市場情緒數據..." +
                "本儀表板整合以下指標：" +
                "• VIX 恐慌指數（CBOE 波動率）" +
                "• S&P 500 / Nasdaq 日漲跌幅" +
                "• VIX 動能（急升=恐懼加速）" +
                "• 市場廣度代理（多指數同向程度）" +
                "合成 Fear & Greed Score：" +
                "  0 ~ 25  = 🟣 極度恐懼（歷史上往往是買點）" +
                "  25 ~ 40 = 🔵 恐懼" +
                "  40 ~ 60 = ⬜ 中立" +
                "  60 ~ 75 = 💚 貪婪" +
                "  75 ~100 = 🔥 極度貪婪（市場過熱警戒）";
            pnlDetail.Controls.Add(txtSentimentDetail);
            pnlDetail.Controls.Add(new Label
            {
                Text = "📊 詳細分析 + VIX解讀 + 倉位邏輯",
                Dock = DockStyle.Top,
                Height = 35,
                ForeColor = Color.Gold,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold)
            });
            split7.Panel2.Controls.Add(pnlDetail);
        }

        private Label MakeSentimentLabel(string title, string val, Color valColor)
        {
            var pnl = new Panel { Dock = DockStyle.Top, Height = 38, BackColor = Color.Transparent };
            var lbl1 = new Label
            {
                Text = title + ":",
                Dock = DockStyle.Left,
                Width = 130,
                ForeColor = Color.LightGray,
                Font = new Font("Segoe UI", 10.5F),
                TextAlign = ContentAlignment.MiddleLeft
            };
            var lbl2 = new Label
            {
                Text = val,
                Dock = DockStyle.Fill,
                ForeColor = valColor,
                Font = new Font("Consolas", 14F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnl.Controls.Add(lbl2); pnl.Controls.Add(lbl1);
            // 回傳值標籤，外部可透過 pnl.Controls[0] 存取
            // 但這樣不方便，直接存到欄位
            return lbl2;
        }

        private void DrawFearGreedBar(Graphics g, int w, int h)
        {
            if (w <= 0 || h <= 0) return;
            // 漸層背景：紫→藍→白→綠→紅
            using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(
                new Rectangle(0, 0, w, h), Color.Purple, Color.Red, 0F))
            {
                var blend = new System.Drawing.Drawing2D.ColorBlend(5);
                blend.Colors = new[] { Color.Purple, Color.DeepSkyBlue, Color.White, Color.LightGreen, Color.OrangeRed };
                blend.Positions = new float[] { 0f, 0.25f, 0.5f, 0.7f, 1f };
                brush.InterpolationColors = blend;
                g.FillRectangle(brush, 0, 0, w, h);
            }
            if (_lastSentiment != null && _lastSentiment.FearGreedScore > 0)
            {
                float x = (float)(_lastSentiment.FearGreedScore / 100.0 * w);
                using (var pen = new Pen(Color.White, 3))
                { g.DrawLine(pen, x, 0, x, h); }
                g.DrawString($"{_lastSentiment.FearGreedScore:F0}", new Font("Consolas", 9F, FontStyle.Bold),
                    Brushes.Black, x + 3, 2);
            }
        }

        // ── 分頁6: 股票篩選器 ────────────────────────────────────────────────
        private void BuildTab6(Panel t)
        {
            var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 320, BackColor = Color.Transparent, SplitterWidth = 8, Orientation = Orientation.Vertical };
            t.Controls.Add(split);

            // 左側: 條件設定
            var pnlCrit = new Panel { Dock = DockStyle.Fill, Padding = new Padding(18), BackColor = Color.FromArgb(28, 30, 38), AutoScroll = true };
            var titleLbl = new Label { Text = "🔍 篩選條件", Dock = DockStyle.Top, Height = 38, ForeColor = Color.White, Font = new Font("Segoe UI", 13F, FontStyle.Bold) };

            // 股票清單
            var lblTickers = new Label { Text = "股票代號清單 (逗號分隔):", Dock = DockStyle.Top, Height = 22, ForeColor = Color.LightGray, Font = new Font("Segoe UI", 10F) };
            txtScreenerTickers = new TextBox { Text = "2330.TW, 2317.TW, 2454.TW, 2308.TW, 2382.TW, NVDA, AAPL, MSFT, TSLA, AMZN", Multiline = true, Height = 60, Dock = DockStyle.Top, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White, Font = new Font("Segoe UI", 10F), BorderStyle = BorderStyle.None, WordWrap = true };

            // RSI 範圍
            var pnlRsi = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.Transparent };
            chkRsiFilter = new CheckBox { Text = "RSI 範圍篩選", Dock = DockStyle.Top, Height = 24, ForeColor = Color.LightCyan, Font = new Font("Segoe UI", 10F) };
            var pnlRsiRange = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 28, BackColor = Color.Transparent };
            numRsiMin = new NumericUpDown { Minimum = 0, Maximum = 100, Value = 0, Width = 68, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White, BorderStyle = BorderStyle.None };
            numRsiMax = new NumericUpDown { Minimum = 0, Maximum = 100, Value = 35, Width = 68, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White, BorderStyle = BorderStyle.None };
            var lblTo = new Label { Text = " ～ ", Width = 30, ForeColor = Color.Gray, TextAlign = ContentAlignment.MiddleCenter };
            pnlRsiRange.Controls.AddRange(new Control[] { numRsiMin, lblTo, numRsiMax });
            pnlRsi.Controls.Add(pnlRsiRange); pnlRsi.Controls.Add(chkRsiFilter);

            // 勾選條件
            chkMacdPos = new CheckBox { Text = "MACD 柱 > 0（動能偏多）", Dock = DockStyle.Top, Height = 26, ForeColor = Color.LightGray, Font = new Font("Segoe UI", 10F) };
            chkMacdCross = new CheckBox { Text = "MACD 黃金交叉（近期翻正）", Dock = DockStyle.Top, Height = 26, ForeColor = Color.LightGray, Font = new Font("Segoe UI", 10F) };
            chkAboveEma50 = new CheckBox { Text = "收盤 > EMA50（短期多頭）", Dock = DockStyle.Top, Height = 26, ForeColor = Color.LightGray, Font = new Font("Segoe UI", 10F) };
            chkAboveEma200 = new CheckBox { Text = "收盤 > EMA200（長期多頭）", Dock = DockStyle.Top, Height = 26, ForeColor = Color.LightGray, Font = new Font("Segoe UI", 10F) };
            chkBBBreakout = new CheckBox { Text = "突破布林上軌（強勢突破）", Dock = DockStyle.Top, Height = 26, ForeColor = Color.LightGray, Font = new Font("Segoe UI", 10F) };
            chkBBOversold = new CheckBox { Text = "跌破布林下軌（超賣反彈）", Dock = DockStyle.Top, Height = 26, ForeColor = Color.LightGray, Font = new Font("Segoe UI", 10F) };

            // 量能倍率
            var pnlVol = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 28, BackColor = Color.Transparent };
            var lblVol = new Label { Text = "5日均量/20日均量 ≥", Width = 155, ForeColor = Color.LightGray, Font = new Font("Segoe UI", 10F), TextAlign = ContentAlignment.MiddleLeft };
            numVolRatio = new NumericUpDown { Minimum = 0, Maximum = 10, Value = 0, DecimalPlaces = 1, Increment = 0.5M, Width = 68, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White, BorderStyle = BorderStyle.None };
            pnlVol.Controls.AddRange(new Control[] { lblVol, numVolRatio });

            btnRunScreener = MakeButton("🔍 執行篩選", Color.FromArgb(0, 140, 180), BtnRunScreener_Click);
            btnRunScreener.Dock = DockStyle.Top; btnRunScreener.Height = 48; btnRunScreener.Margin = new Padding(0, 12, 0, 0);

            // 由下往上 Add (Dock=Top 後加的在上方)
            pnlCrit.Controls.Add(btnRunScreener);
            pnlCrit.Controls.Add(pnlVol);
            pnlCrit.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 8 });
            pnlCrit.Controls.Add(chkBBOversold); pnlCrit.Controls.Add(chkBBBreakout);
            pnlCrit.Controls.Add(chkAboveEma200); pnlCrit.Controls.Add(chkAboveEma50);
            pnlCrit.Controls.Add(chkMacdCross); pnlCrit.Controls.Add(chkMacdPos);
            pnlCrit.Controls.Add(pnlRsi);
            pnlCrit.Controls.Add(new Panel { Dock = DockStyle.Top, Height = 8 });
            pnlCrit.Controls.Add(txtScreenerTickers); pnlCrit.Controls.Add(lblTickers);
            pnlCrit.Controls.Add(titleLbl);
            split.Panel1.Controls.Add(pnlCrit);

            // 右側: 結果
            var pnlResult = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12), BackColor = Color.FromArgb(28, 30, 38) };
            txtScreenerLog = MakeTextBox(Color.LightGray, new Font("Consolas", 10F));
            txtScreenerLog.Height = 80; txtScreenerLog.Dock = DockStyle.Bottom;
            dgvScreener = BuildDGV();
            pnlResult.Controls.Add(dgvScreener);
            pnlResult.Controls.Add(txtScreenerLog);
            pnlResult.Controls.Add(new Label { Text = "📊 篩選結果 (按符合條件數排序)", Dock = DockStyle.Top, Height = 35, ForeColor = Color.White, Font = new Font("Segoe UI", 12F, FontStyle.Bold) });
            split.Panel2.Controls.Add(pnlResult);
        }

        // ── 輔助建構方法 ─────────────────────────────────────────────────────
        private Button MakeButton(string text, Color bg, EventHandler handler)
        {
            var btn = new Button { Text = text, Width = 160, Height = 42, BackColor = bg, ForeColor = Color.White, Font = new Font("Segoe UI", 10.5F, FontStyle.Bold), FlatStyle = FlatStyle.Flat };
            btn.FlatAppearance.BorderSize = 0; btn.Click += handler; return btn;
        }

        private DataGridView BuildDGV()
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.FromArgb(18, 20, 28),
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                EnableHeadersVisualStyles = false,
                AllowUserToAddRows = false,
                ReadOnly = true,
                GridColor = Color.FromArgb(35, 40, 50),
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                RowTemplate = { Height = 38 }
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(25, 27, 35);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.LightGray;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            dgv.ColumnHeadersHeight = 42; dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dgv.DefaultCellStyle.BackColor = Color.FromArgb(18, 20, 28); dgv.DefaultCellStyle.ForeColor = Color.White;
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(40, 80, 150); dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.RowsDefaultCellStyle.BackColor = Color.FromArgb(18, 20, 28); dgv.RowsDefaultCellStyle.ForeColor = Color.White;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(22, 24, 32); dgv.AlternatingRowsDefaultCellStyle.ForeColor = Color.White;
            return dgv;
        }

        private TextBox MakeTextBox(Color fore, Font font)
        {
            return new TextBox { Multiline = true, ReadOnly = true, Dock = DockStyle.Fill, BackColor = Color.FromArgb(18, 20, 28), ForeColor = fore, Font = font, BorderStyle = BorderStyle.None, ScrollBars = ScrollBars.Vertical };
        }

        private Panel MakeLabeledPanel(string title, Control content, Color titleColor, Color bg = default)
        {
            var p = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15, 15, 15, 8), BackColor = bg == default ? Color.FromArgb(28, 30, 38) : bg };
            p.Controls.Add(content);
            p.Controls.Add(new Label { Text = title, Dock = DockStyle.Top, Height = 36, ForeColor = titleColor, Font = new Font("Segoe UI", 11F, FontStyle.Bold) });
            return p;
        }

        private Control CreateInputWrapper(Control input, string labelText)
        {
            var wrapper = new Panel { Width = input.Width + 28, Height = 44, Margin = new Padding(8, 0, 8, 0), BackColor = Color.Transparent };
            var lbl = new Label { Text = labelText, ForeColor = Color.Gray, Font = new Font("Segoe UI", 9F), Dock = DockStyle.Top, Height = 17 };
            var bg = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(45, 48, 55), Padding = new Padding(8, 5, 8, 4) };
            input.Dock = DockStyle.Fill; bg.Controls.Add(input); wrapper.Controls.Add(bg); wrapper.Controls.Add(lbl);
            return wrapper;
        }

        private Control CreateComboWrapper(ComboBox cb, string labelText)
        {
            var wrapper = new Panel { Width = cb.Width + 18, Height = 44, Margin = new Padding(8, 0, 8, 0), BackColor = Color.Transparent };
            var lbl = new Label { Text = labelText, ForeColor = Color.Gray, Font = new Font("Segoe UI", 9F), Dock = DockStyle.Top, Height = 17 };
            cb.Dock = DockStyle.Fill; wrapper.Controls.Add(cb); wrapper.Controls.Add(lbl);
            return wrapper;
        }

        private Label CreateStatLabel(Control parent, string title, string val, int x)
        {
            var panel = new Panel { Size = new Size(128, 80), Location = new Point(x, 5), BackColor = Color.Transparent };
            var l1 = new Label { Text = title, Dock = DockStyle.Top, Height = 22, ForeColor = Color.Gray, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 9F) };
            var l2 = new Label { Text = val, Dock = DockStyle.Fill, ForeColor = Color.FromArgb(0, 255, 150), TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Consolas", 15F, FontStyle.Bold) };
            panel.Controls.Add(l2); panel.Controls.Add(l1); parent.Controls.Add(panel); return l2;
        }

        // ── 初始化 ────────────────────────────────────────────────────────────
        private async void InitDataAsync()
        {
            string saved = ApiKeyManager.Load();
            if (!string.IsNullOrEmpty(saved)) txtApiKey.Text = saved;
            UpdateCachedTickers();
            await UpdateTickerInfoUI();
            AutoLoadCachedData();
        }

        private void UpdateCachedTickers()
        {
            Task.Run(() =>
            {
                var tickers = ExcelCacheManager.GetCachedTickers();
                if (tickers == null || tickers.Count == 0) return;
                this.Invoke(new Action(() =>
                {
                    if (txtTicker.Focused) return;
                    var cur = txtTicker.Items.Cast<string>().ToList();
                    if (tickers.Count != cur.Count || !tickers.All(cur.Contains))
                    {
                        string ct = txtTicker.Text;
                        txtTicker.BeginUpdate(); txtTicker.Items.Clear();
                        txtTicker.Items.AddRange(tickers.ToArray<object>());
                        txtTicker.EndUpdate();
                        if (txtTicker.Text != ct) txtTicker.Text = ct;
                    }
                }));
            });
        }

        private void AutoLoadCachedData()
        {
            string ticker = txtTicker.Text.Trim().ToUpper();
            if (string.IsNullOrEmpty(ticker)) return;
            var cached = ExcelCacheManager.LoadData(ticker);
            if (cached != null && cached.Count > 0)
            {
                IndicatorEngine.CalculateAll(cached);
                SupportResistanceEngine.CalculatePivots(cached);
                historyData = cached.TakeLast((int)numDays.Value).ToList();
                RefreshHistoryGrid();
                historyChart.UpdateData(historyData, ticker, _currentCompanyName, _currentPE, _currentYield);
                lblStatus.Text = $"載入 {ticker} 本地快取完成。";
            }
            else
            {
                lblStatus.Text = $"尚無 {ticker} 快取，請執行分析。";
            }
        }

        private void SwitchToLiveTab()
        {
            string ticker = txtTicker.Text.Trim().ToUpper();

            // ── 股票不同時清空所有即時狀態，避免舊數據汙染 ──
            if (ticker != _currentLiveTicker)
            {
                currentLiveData.Clear();
                _currentLiveTicker = ticker;
                _lastTickTime = "";
                _lastTickPrice = -1;
                _lastAccumulatedVolume = -1;
                lblLivePrice.Text = "$ --";
                lblLiveChange.Text = "+0.00%";
                lblLiveSentiment.Text = "⚪ 等待數據...";
                lblLiveHeartbeat.Text = "切換至 " + ticker;
                lblLiveHeartbeat.ForeColor = Color.Gold;
                orderBookPanel.ClearData($"已切換至 {ticker}，等待即時報價...");
            }

            // 嘗試從今日快取載入（同一支股票才讀）
            var cached = ExcelIntradayCacheManager.LoadTodayData(ticker);
            if (cached.Count > 0)
            {
                currentLiveData = cached;
                _lastTickTime = currentLiveData.Last().Date.ToString("yyyy-MM-dd HH:mm:ss");
                _lastTickPrice = currentLiveData.Last().Close;
                _lastAccumulatedVolume = currentLiveData.Last().Volume;
                IndicatorEngine.CalculateAll(currentLiveData);
                SupportResistanceEngine.CalculatePivots(currentLiveData);
            }

            liveChart.UpdateData(currentLiveData, ticker + " [即時]", _currentCompanyName, _currentPE, _currentYield);
        }

        private async Task UpdateTickerInfoUI()
        {
            string ticker = txtTicker.Text.Trim().ToUpper();
            if (string.IsNullOrEmpty(ticker)) return;
            bool isUS = cbMarketType.SelectedIndex == 1;
            try
            {
                _currentCompanyName = await MarketInfoService.GetCompanyName(ticker, isUS);

                // ── ETF 自動偵測 ──────────────────────────────────────────
                bool isEtf = await EtfService.IsEtfAsync(ticker);
                if (isEtf)
                {
                    _etfInfo = await EtfService.FetchEtfInfo(ticker);
                    UpdateEtfCard(_etfInfo);
                    // ETF 使用殖利率，PE 通常無意義
                    _currentPE = 0;
                    _currentYield = _etfInfo?.DividendYield ?? 0;
                }
                else
                {
                    _etfInfo = null;
                    if (pnlEtfCard != null) pnlEtfCard.Visible = false;
                    if (lblEtfBadge != null) lblEtfBadge.Visible = false;

                    var (pe, yield, source) = await YahooDataService.FetchFundamentals(ticker);
                    _currentPE = pe;
                    _currentYield = yield;
                    if (source == "Unavailable")
                    {
                        bool noAvKey = !AlphaVantageKeyManager.HasKey;
                        string hint = noAvKey ? "  建議在「提示詞」頁填入 Alpha Vantage Key" : "";
                        ShowToast($"⚠️ {ticker} 基本面無法取得（所有來源失敗）{hint}", Color.FromArgb(100, 70, 20));
                    }
                    else if (source != "TWSE" && source != "TPEX")
                        lblStatus.Text = $"{ticker} 基本面來源: {source}  P/E {(pe > 0 ? pe.ToString("F2") : "N/A")}  殖利率 {(yield > 0 ? yield.ToString("P2") : "N/A")}";
                }

                _lastInfoTicker = ticker;

                // ── 財報日（美股個股才查）────────────────────────────────
                if (isUS && !isEtf && lblEarningsDate != null)
                {
                    _ = Task.Run(async () =>
                    {
                        var (nextDate, est, act, surp, quarter) = await EarningsService.FetchEarningsAsync(ticker);
                        this.Invoke(new Action(() =>
                        {
                            if (nextDate.HasValue)
                            {
                                int daysLeft = (int)(nextDate.Value.Date - DateTime.Today).TotalDays;
                                string urgency = daysLeft <= 7 ? "🔴" : daysLeft <= 21 ? "🟡" : "🟢";
                                string estStr = est != 0 ? $"EPS 預估 ${est:F2}" : "";
                                lblEarningsDate.Text = $"{urgency} 財報 {nextDate.Value:MM/dd} ({daysLeft}天)  {estStr}";
                                lblEarningsDate.Visible = true;
                                lblEarningsDate.ForeColor = daysLeft <= 7 ? Color.OrangeRed : Color.Orange;
                            }
                            else { lblEarningsDate.Visible = false; }
                        }));
                    });
                }
                else if (lblEarningsDate != null) lblEarningsDate.Visible = false;

                if (historyData.Count > 0) historyChart.UpdateData(historyData, ticker, _currentCompanyName, _currentPE, _currentYield);
                if (currentLiveData.Count > 0) liveChart.UpdateData(currentLiveData, ticker, _currentCompanyName, _currentPE, _currentYield);
            }
            catch (Exception ex) { AppLogger.Log("UpdateTickerInfoUI 失敗", ex); }
        }

        // ── 更新歷史表格 ─────────────────────────────────────────────────────
        private void RefreshHistoryGrid()
        {
            dgvHistory.DataSource = historyData.Select(d => new
            {
                日期 = d.Date.ToString("yyyy-MM-dd"),
                收盤 = d.Close.ToString("F2"),
                支撐 = d.SupportLevel.ToString("F2"),
                壓力 = d.ResistanceLevel.ToString("F2"),
                RSI = d.RSI.ToString("F1"),
                MACD = d.MACD.ToString("F3"),
                Signal = d.MACD_Signal.ToString("F3"),
                MACD柱 = d.MACD_Hist.ToString("F3"),
                BBW = d.BB_Width.ToString("F1") + "%",
                形態 = d.Pattern,
                AI建議 = d.AgentAction
            }).ToList();

            foreach (DataGridViewRow row in dgvHistory.Rows)
            {
                string action = row.Cells["AI建議"].Value?.ToString() ?? "";
                if (action == "Buy") row.DefaultCellStyle.ForeColor = Color.SpringGreen;
                else if (action == "Sell") row.DefaultCellStyle.ForeColor = Color.LightCoral;
            }
        }

        // ── 執行分析（支援取消 + 真實進度條）────────────────────────────────
        private async void BtnRunAnalysis_Click(object sender, EventArgs e)
        {
            // ← 取消邏輯：若分析中則取消，否則啟動
            if (_analysisCts != null)
            {
                _analysisCts.Cancel();
                btnRunAnalysis.Text = "✨ 執行分析";
                return;
            }

            currentApiKey = txtApiKey.Text.Trim();
            if (string.IsNullOrEmpty(currentApiKey))
            {
                ShowToast("⚠️ 請先輸入 API Key", Color.FromArgb(120, 50, 0));
                return;
            }

            _analysisCts = new CancellationTokenSource();
            var ct = _analysisCts.Token;
            btnRunAnalysis.Text = "✕ 取消分析";
            progressBar.Visible = true;
            progressBar.Value = 0;
            txtNewsLog.Clear();

            try
            {
                string ticker = txtTicker.Text.Trim().ToUpper();
                int days = (int)numDays.Value;
                if (ticker != _lastInfoTicker || string.IsNullOrEmpty(_currentCompanyName))
                    await UpdateTickerInfoUI();

                // ── 階段 1/4：新聞（20%）────────────────────────────────────
                string kw = !string.IsNullOrEmpty(_currentCompanyName) ? _currentCompanyName : ticker;
                lblStatus.Text = $"[1/4] 抓取 {kw} 新聞...";
                progressBar.Value = 5;
                currentNews = await NewsService.FetchCnaNews(kw);
                txtNewsLog.Text = currentNews.Count > 0 ? string.Join("\r\n\r\n", currentNews) : "查無相關新聞。";
                progressBar.Value = 20;
                ct.ThrowIfCancellationRequested();

                // ── 階段 2/4：歷史資料（40%）────────────────────────────────
                lblStatus.Text = "[2/4] 讀取本地快取...";
                var cached = ExcelCacheManager.LoadData(ticker);
                DateTime? lastDate = cached.Count > 0 ? cached.Max(d => d.Date) : (DateTime?)null;

                List<MarketData> fetched;
                if (lastDate.HasValue && lastDate.Value.Date >= DateTime.Now.AddDays(-5))
                {
                    lblStatus.Text = "[2/4] 補齊近期數據...";
                    fetched = await YahooDataService.FetchYahooByDate(ticker, lastDate.Value.AddDays(-2), DateTime.Now, ct);
                }
                else
                {
                    lblStatus.Text = "[2/4] 下載3年歷史數據...";
                    fetched = await YahooDataService.FetchYahoo(ticker, "1d", "3y", ct);
                }

                var merged = cached.ToDictionary(d => d.Date.Date);
                foreach (var d in fetched) merged[d.Date.Date] = d;
                historyData = merged.Values.OrderBy(d => d.Date).ToList();
                ExcelCacheManager.SaveData(ticker, historyData);
                IndicatorEngine.CalculateAll(historyData);
                SupportResistanceEngine.CalculatePivots(historyData);
                FeedbackEngine.EvaluatePastActions(ticker, historyData);
                string rlLessons = FeedbackEngine.GetRecentLessons(ticker);
                historyData = historyData.TakeLast(days).ToList();
                progressBar.Value = 40;
                ct.ThrowIfCancellationRequested();

                // ── 階段 3/4：AI 分析（85%）─────────────────────────────────
                string fundStr;
                if (_etfInfo != null && _etfInfo.IsEtf)
                    fundStr = $"【ETF基本面】{_etfInfo.Category}  AUM: {(_etfInfo.TotalAssets >= 1e9 ? $"{_etfInfo.TotalAssets / 1e9:F1}B" : "N/A")}  費率: {_etfInfo.ExpenseRatio:P2}  殖利率: {_etfInfo.DividendYield:P2}\n" +
                              (_etfInfo.TopHoldings.Count > 0
                                  ? "前五大持股: " + string.Join(", ", _etfInfo.TopHoldings.Take(5).Select(h => $"{h.Name}({h.Pct:P0})")) + "\n"
                                  : "");
                else
                    fundStr = _currentPE > 0
                        ? $"【基本面】本益比(P/E): {_currentPE:F2}，殖利率: {_currentYield:P2}\n"
                        : "【基本面】P/E: N/A，殖利率: N/A\n";
                string newsCtx = fundStr + string.Join("\n", currentNews);
                string sentimentCtx = "";
                if (_lastSentiment != null && !_lastSentiment.IsStale)
                    sentimentCtx = $"\n【市場整體情緒警示】VIX={_lastSentiment.VIX:F2} ({_lastSentiment.VIXLevel}) | " +
                                   $"Fear&Greed={_lastSentiment.FearGreedScore:F0}/100 ({_lastSentiment.FearGreedLabel}) | " +
                                   $"{_lastSentiment.AiPositionAdvice}\n" +
                                   $"👉 請根據以上整體市場情緒調整操作積極程度。\n";
                string finalPrompt = txtSystemPrompt.Text + "\n" + rlLessons + sentimentCtx;
                lblStatus.Text = $"[3/4] AI 多模態分析 ({LlmConfig.CurrentModel})...";
                progressBar.Value = 45;

                // ── 台股：平行抓取三大法人 + 融資融券（真實籌碼面資料）────────
                InstitutionalData instData = null;
                MarginData marginData = null;
                bool isTwStock = ticker.EndsWith(".TW") || ticker.EndsWith(".TWO");
                if (isTwStock)
                {
                    lblStatus.Text = $"[3/4] 抓取三大法人 + 融資融券籌碼...";
                    var instTask   = InstitutionalService.FetchAsync(ticker);
                    var marginTask = MarginTradingService.FetchAsync(ticker);
                    await Task.WhenAll(instTask, marginTask);
                    instData   = instTask.Result;
                    marginData = marginTask.Result;
                }
                lblStatus.Text = $"[3/4] AI 多模態分析 ({LlmConfig.CurrentModel})...";

                await AgentEngine.RunAlphaDebate(historyData, newsCtx, currentApiKey, finalPrompt,
                    (log) => this.Invoke(new Action(() =>
                        txtDebateLog.Text = (!string.IsNullOrEmpty(rlLessons) ? rlLessons + "\n\n" : "") + log)),
                    _lastMTFSignal, ct,
                    institutional: instData,
                    margin: marginData);
                progressBar.Value = 85;
                ct.ThrowIfCancellationRequested();

                // ── 階段 4/4：回測統計（100%）───────────────────────────────
                lblStatus.Text = "[4/4] 回測計算...";
                var bt = BacktestEngine.RunBacktest(historyData);
                lblTotalReturn.Text = $"{bt.TotalReturn:P2}";
                lblTotalReturn.ForeColor = bt.TotalReturn >= 0 ? Color.SpringGreen : Color.LightCoral;
                lblWinRate.Text = $"{bt.WinRate:P1}";
                lblMDD.Text = $"{bt.MaxDrawdown:P2}";
                lblTradeCount.Text = $"{bt.TradeCount}";
                lblSharpe.Text = $"{bt.SharpeRatio:F2}";
                lblSharpe.ForeColor = bt.SharpeRatio >= 1 ? Color.SpringGreen : (bt.SharpeRatio >= 0 ? Color.Gold : Color.LightCoral);
                lblSortino.Text = $"{bt.SortinoRatio:F2}";
                lblSortino.ForeColor = bt.SortinoRatio >= 1 ? Color.SpringGreen : (bt.SortinoRatio >= 0 ? Color.Gold : Color.LightCoral);
                lblKelly.Text = $"{bt.KellyFraction:P1}";
                lblKelly.ForeColor = Color.Gold;
                RefreshHistoryGrid();
                historyChart.UpdateData(historyData, ticker, _currentCompanyName, _currentPE, _currentYield);
                progressBar.Value = 100;
                lblStatus.Text = $"✅ 分析完成。Sharpe={bt.SharpeRatio:F2}  Kelly={bt.KellyFraction:P1}  勝率={bt.WinRate:P1}";
                ShowToast($"✅ {ticker} 分析完成  Sharpe {bt.SharpeRatio:F2}  勝率 {bt.WinRate:P1}", Color.FromArgb(20, 90, 40));
            }
            catch (OperationCanceledException)
            {
                lblStatus.Text = "⏹ 分析已取消";
                ShowToast("⏹ 分析已取消", Color.FromArgb(60, 60, 60));
            }
            catch (Exception ex)
            {
                // ← 修正：改用 Toast，不用阻斷式 MessageBox
                AppLogger.Log("BtnRunAnalysis_Click 失敗", ex);
                lblStatus.Text = $"❌ 分析失敗：{ex.Message}";
                ShowToast($"❌ 分析失敗：{ex.Message.Substring(0, Math.Min(ex.Message.Length, 60))}", Color.FromArgb(120, 30, 20));
            }
            finally
            {
                _analysisCts?.Dispose();
                _analysisCts = null;
                btnRunAnalysis.Text = "✨ 執行分析";
                progressBar.Visible = false;
                progressBar.Value = 0;
            }
        }

        // ── 多時間框架分析 ────────────────────────────────────────────────────
        private async void BtnMTF_Click(object sender, EventArgs e)
        {
            string ticker = txtTicker.Text.Trim().ToUpper();
            if (string.IsNullOrEmpty(ticker)) return;
            btnMTF.Enabled = false; btnMTF.Text = "🔀 分析中...";
            txtMTFResult.Text = "正在拉取週線/日線/小時線數據，請稍候...\r\n";
            try
            {
                _lastMTFSignal = await MultiTimeframeEngine.AnalyzeAsync(ticker, _analysisCts?.Token ?? CancellationToken.None);
                txtMTFResult.Text = _lastMTFSignal.AlignmentSummary +
                    $"\r\n\r\n═══ 週線 ═══\r\nRSI={_lastMTFSignal.Weekly_RSI:F1}  MACD柱={_lastMTFSignal.Weekly_MACD_Hist:F3}  趨勢={_lastMTFSignal.Weekly_Trend}  形態={_lastMTFSignal.Weekly_Pattern}" +
                    $"\r\n═══ 日線 ═══\r\nRSI={_lastMTFSignal.Daily_RSI:F1}  MACD柱={_lastMTFSignal.Daily_MACD_Hist:F3}  趨勢={_lastMTFSignal.Daily_Trend}  形態={_lastMTFSignal.Daily_Pattern}" +
                    $"\r\n═══ 小時線 ═══\r\nRSI={_lastMTFSignal.Hourly_RSI:F1}  MACD柱={_lastMTFSignal.Hourly_MACD_Hist:F3}  趨勢={_lastMTFSignal.Hourly_Trend}  形態={_lastMTFSignal.Hourly_Pattern}" +
                    $"\r\n\r\n💡 下次「執行分析」時，多時間框架信號將自動注入 AI 提示詞。";
                lblStatus.Text = $"多時間框架分析完成。共振分數={_lastMTFSignal.AlignmentScore}";
            }
            catch (Exception ex) { AppLogger.Log("BtnMTF_Click 失敗", ex); txtMTFResult.Text = $"失敗: {ex.Message}"; }
            finally { btnMTF.Enabled = true; btnMTF.Text = "🔀 多時間框架分析 (週/日/時)"; }
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 停止並釋放計時器、取消分析、釋放Semaphore，避免關閉時背景工作仍在執行
            try
            {
                // 取消正在進行的分析
                try { _analysisCts?.Cancel(); } catch { }
                try { _analysisCts?.Dispose(); } catch { }
                _analysisCts = null;

                // Timers
                try { liveTimer?.Stop(); liveTimer?.Dispose(); } catch { }
                liveTimer = null;
                try { alertTimer?.Stop(); alertTimer?.Dispose(); } catch { }
                alertTimer = null;
                try { cacheTimer?.Stop(); cacheTimer?.Dispose(); } catch { }
                cacheTimer = null;
                try { toastTimer?.Stop(); toastTimer?.Dispose(); } catch { }
                toastTimer = null;

                // SemaphoreSlim
                try { netLock?.Dispose(); } catch { }
                netLock = null;

                // 關閉浮動聊天視窗（若開啟）
                try
                {
                    if (chatFormInstance != null && !chatFormInstance.IsDisposed)
                    {
                        chatFormInstance.Close();
                        chatFormInstance.Dispose();
                    }
                }
                catch { }
                chatFormInstance = null;
            }
            finally
            {
                base.OnFormClosing(e);
            }
        }
        // ── 當沖 Live ─────────────────────────────────────────────────────────
        private async Task RefreshLive()
        {
            if (!await netLock.WaitAsync(0)) return;
            string ticker = txtTicker.Text.Trim().ToUpper();

            // 💡 整合優化：ticker 已切換，立刻清空舊數據再繼續
            if (ticker != _currentLiveTicker)
            {
                currentLiveData.Clear();
                _currentLiveTicker = ticker;
                _lastTickTime = "";
                _lastTickPrice = -1;
                _lastAccumulatedVolume = -1;
                liveChart.UpdateData(currentLiveData, ticker + " [即時動能]", _currentCompanyName, _currentPE, _currentYield);
            }

            if (ticker != _lastInfoTicker && !string.IsNullOrEmpty(ticker))
            {
                await UpdateTickerInfoUI();
            }

            bool isUSMarket = cbMarketType.SelectedIndex == 1;
            try
            {
                lblLiveHeartbeat.Text = "連線中..."; lblLiveHeartbeat.ForeColor = Color.Gold;

                if (isUSMarket)
                {
                    orderBookPanel.ClearData("美股市場不提供五檔掛單資料");
                    var data = await YahooDataService.FetchYahoo(ticker, "1m", "1d");

                    // 💡 新增判斷：美股當資料的成交量是 0 時，過濾掉不加入線圖
                    data = data.Where(x => x.Volume > 0).ToList();

                    if (data.Count > 0)
                    {
                        IndicatorEngine.CalculateAll(data);
                        SupportResistanceEngine.CalculatePivots(data);
                        currentLiveData = data;
                        liveChart.UpdateData(data, ticker + " [US 即時動能]", _currentCompanyName, _currentPE, _currentYield);
                        var last = data.Last(); var first = data.First();
                        double diff = last.Close - first.Close;
                        lblLivePrice.Text = string.Concat("$", last.Close);
                        lblLivePrice.ForeColor = diff >= 0 ? Color.SpringGreen : Color.LightCoral; // 與變動一致顯色
                        lblLivePrice.Visible = true;
                        lblLiveChange.Text = $"{(diff >= 0 ? "+" : "")}{diff:F2} ({(diff / first.Close):P2})";
                        lblLiveChange.ForeColor = diff >= 0 ? Color.SpringGreen : Color.LightCoral;
                        lblLiveHeartbeat.Text = $"Yahoo Sync: {DateTime.Now:HH:mm:ss}"; lblLiveHeartbeat.ForeColor = Color.SpringGreen;
                        UpdateLiveSentiment(last);
                    }
                }
                else
                {
                    bool isTaiwanStock = ticker.EndsWith(".TW") || ticker.EndsWith(".TWO");
                    if (!isTaiwanStock) { orderBookPanel.ClearData("台股請輸入 .TW 或 .TWO 結尾"); lblLiveHeartbeat.Text = "格式錯誤"; return; }

                    var twseTick = await TwseApiService.FetchRealtime(ticker);
                    if (twseTick != null && twseTick.Price > 0)
                    {
                        // 💡 新增判斷：當台股資料的單筆成交量 (TradeVolume) 是 0 時，就 return 不加入線圖
                        if (twseTick.TradeVolume <= 0) return;

                        if (!string.IsNullOrEmpty(twseTick.CompanyName) && twseTick.CompanyName != "-")
                        {
                            _currentCompanyName = twseTick.CompanyName;
                        }

                        orderBookPanel.UpdateData(twseTick);
                        bool isDuplicateTrade = (twseTick.Price == _lastTickPrice && twseTick.Volume == _lastAccumulatedVolume);

                        if (!isDuplicateTrade)
                        {
                            currentLiveData.Add(new MarketData
                            {
                                Date = DateTime.Parse(twseTick.Time),
                                Open = twseTick.Open > 0 ? twseTick.Open : twseTick.Price,
                                High = twseTick.High > 0 ? twseTick.High : twseTick.Price,
                                Low = twseTick.Low > 0 ? twseTick.Low : twseTick.Price,
                                Close = twseTick.Price,
                                Volume = twseTick.TradeVolume * 1000
                            });
                            // ← 修正：RemoveAt(0) 是 O(n)，改為批次刪除 500 筆（均攤 O(1)）
                            if (currentLiveData.Count > LiveDataMaxSize)
                                currentLiveData.RemoveRange(0, LiveDataTrimBatch);
                            _lastTickTime = twseTick.Time; _lastTickPrice = twseTick.Price; _lastAccumulatedVolume = twseTick.Volume;

                            ExcelIntradayCacheManager.SaveData(ticker, currentLiveData);
                            if (currentLiveData.Count > 0)
                            {
                                IndicatorEngine.CalculateAll(currentLiveData);
                                SupportResistanceEngine.CalculatePivots(currentLiveData);
                                liveChart.UpdateData(currentLiveData, ticker + " [TWSE Tick 走勢]", _currentCompanyName, _currentPE, _currentYield);
                                UpdateLiveSentiment(currentLiveData.Last());
                            }
                        }
                        double diff = twseTick.Price - twseTick.YesterdayClose;
                        lblLivePrice.Text = $"$ {twseTick.Price:F2}";
                        lblLivePrice.ForeColor = diff >= 0 ? Color.SpringGreen : Color.LightCoral;
                        lblLivePrice.Visible = true;
                        lblLiveChange.Text = $"{(diff >= 0 ? "+" : "")}{diff:F2} ({(diff / twseTick.YesterdayClose):P2})";
                        lblLiveChange.ForeColor = diff >= 0 ? Color.SpringGreen : Color.LightCoral;
                        lblLiveHeartbeat.Text = isDuplicateTrade ? $"TWSE 掛單跳動: {twseTick.Time}" : $"TWSE 新成交: {twseTick.Time}";
                        lblLiveHeartbeat.ForeColor = isDuplicateTrade ? Color.LightGray : Color.SpringGreen;

                        // 若您的專案有實作 CheckAlertsForPrice，請保留原本的呼叫
                        CheckAlertsForPrice(ticker, twseTick.Price);
                    }
                    else
                    {
                        orderBookPanel.ClearData("盤後休市或無掛單資料");
                        lblLiveHeartbeat.Text = "無即時數據(可能休市)"; lblLiveHeartbeat.ForeColor = Color.Orange;
                        lblLiveSentiment.Text = "💤 盤後休市中"; lblLiveSentiment.ForeColor = Color.Gray;
                    }
                }
            }
            catch (Exception ex) { AppLogger.Log("RefreshLive 失敗", ex); lblLiveHeartbeat.Text = "連線失敗重試中"; lblLiveHeartbeat.ForeColor = Color.LightCoral; }
            finally { netLock.Release(); }
        }


        private void UpdateLiveSentiment(MarketData last)
        {
            string s = "⚪ 盤整累積中"; Color c = Color.Gray;
            if (currentLiveData.Count > 15)
            {
                if (last.RSI > 75) { s = "🔴 極端超買 (誘多警戒)"; c = Color.LightCoral; }
                else if (last.RSI < 25) { s = "🟢 極端超賣 (尋找底部)"; c = Color.SpringGreen; }
                else if (last.MACD_Hist > 0 && last.RSI > 55) { s = "🔥 多頭動能增溫"; c = Color.Orange; }
                else if (last.MACD_Hist < 0 && last.RSI < 45) { s = "📉 空頭動能發散"; c = Color.DeepSkyBlue; }
                else s = "⚪ 量縮盤整中";
            }
            lblLiveSentiment.Text = s; lblLiveSentiment.ForeColor = c;
        }

        //private async void BtnLiveAiScan_Click(object sender, EventArgs e)
        //{
        //    currentApiKey = txtApiKey.Text.Trim();
        //    if (string.IsNullOrEmpty(currentApiKey)) { MessageBox.Show("請輸入 API Key"); return; }
        //    if (currentLiveData.Count < 5) { MessageBox.Show("即時數據不足，請等待累積。"); return; }
        //    btnLiveAiScan.Enabled = false; btnLiveAiScan.Text = "⚡ 解讀中...";
        //    txtLiveAiDiagnosis.Text = "AI 正在融合即時數據、新聞與 VWAP 分析，請稍候...\r\n\r\n";
        //    try
        //    {
        //        string ticker = txtTicker.Text.Trim().ToUpper();
        //        string kw = !string.IsNullOrEmpty(_currentCompanyName) ? _currentCompanyName : ticker;
        //        if (currentNews.Count == 0) currentNews = await NewsService.FetchCnaNews(kw);
        //        string newsStr = currentNews.Count > 0 ? string.Join("\n", currentNews) : "無重大新聞。";
        //        var td = currentLiveData.TakeLast(30).ToList();
        //        var sb = new StringBuilder("【即時新聞】\n" + newsStr + "\n\n【即時走勢】\nTime,Close,Vol,RSI,MACD,VWAP,Sup,Res\n");
        //        foreach (var d in td) sb.AppendLine($"{d.Date:HH:mm},{d.Close:F2},{d.Volume},{d.RSI:F1},{d.MACD_Hist:F3},{d.VWAP:F2},{d.SupportLevel:F2},{d.ResistanceLevel:F2}");

        //        // 附上 Fib 資訊
        //        if (currentLiveData.Count >= 10)
        //        {
        //            var fibs = FibonacciEngine.Calculate(currentLiveData, Math.Min(120, currentLiveData.Count));
        //            if (fibs.Count > 0) { sb.Append("\n【Fibonacci 回調位】\n"); foreach (var f in fibs) sb.AppendLine($"{f.Label}: {f.Price:F2}"); }
        //        }

        //        await LLMService.StreamChat(currentApiKey, new List<object> { new { role = "system", content = txtLivePrompt.Text }, new { role = "user", content = sb.ToString() } },
        //            (chunk) => this.Invoke(new Action(() => { txtLiveAiDiagnosis.AppendText(chunk); txtLiveAiDiagnosis.ScrollToCaret(); })));
        //    }
        //    catch (Exception ex) { AppLogger.Log("BtnLiveAiScan_Click 失敗", ex); txtLiveAiDiagnosis.Text = $"掃描失敗: {ex.Message}"; }
        //    finally { btnLiveAiScan.Enabled = true; btnLiveAiScan.Text = "⚡ AI 診斷即時動能 + VWAP + Fib"; }
        //}
        private async void BtnLiveAiScan_Click(object sender, EventArgs e)
        {
            currentApiKey = txtApiKey.Text.Trim();
            if (string.IsNullOrEmpty(currentApiKey)) { MessageBox.Show("請輸入 API Key"); return; }
            if (currentLiveData.Count < 5) { MessageBox.Show("即時數據不足，請等待累積。"); return; }

            btnLiveAiScan.Enabled = false; btnLiveAiScan.Text = "⚡ 解讀中...";
            txtLiveAiDiagnosis.Text = "AI 正在融合即時數據、新聞與 VWAP 分析，請稍候...\r\n\r\n";
            // 💡 動態將字體切換為等寬字體 (Consolas)，這樣表格才會完美對齊
            txtLiveAiDiagnosis.Font = new Font("Consolas", 11F);
            txtLiveAiDiagnosis.Text = "資料排版中，準備呼叫 AI...\r\n";
            try
            {
                bool isUS = cbMarketType.SelectedIndex == 1;
                string ticker = txtTicker.Text.Trim().ToUpper();
                string searchKeyword = !string.IsNullOrEmpty(_currentCompanyName) && _currentCompanyName != "-" ? _currentCompanyName : ticker;

                if (currentNews.Count == 0) { currentNews = await NewsService.FetchCnaNews(searchKeyword); }

                string newsStr = currentNews.Count > 0 ? string.Join("\n", currentNews) : "無重大新聞。";

                var td = currentLiveData.TakeLast(30).ToList();
                var sb = new StringBuilder("【即時新聞】\n" + newsStr + "\n\n【即時走勢】\nTime,Close,Vol,RSI,MACD,VWAP,Sup,Res\n");
                foreach (var d in td) sb.AppendLine($"{d.Date:HH:mm},{d.Close:F2},{d.Volume},{d.RSI:F1},{d.MACD_Hist:F3},{d.VWAP:F2},{d.SupportLevel:F2},{d.ResistanceLevel:F2}");

                // 💡 附上 Fib 資訊 (使用您本地實作的 FibonacciEngine)
                if (currentLiveData.Count >= 10)
                {
                    var fibs = FibonacciEngine.Calculate(currentLiveData, Math.Min(120, currentLiveData.Count));
                    if (fibs.Count > 0) { sb.Append("\n【Fibonacci 回調位】\n"); foreach (var f in fibs) sb.AppendLine($"{f.Label}: {f.Price:F2}"); }
                }

                // 💡 核心優化：在 Prompt 中強制要求 AI 自動排版以便觀看
                string formatRule = "\n\n🚨【排版強制要求】：請使用「重點標題」、「條列式 (Bullet Points)」、以及「適當的 Emoji 符號」來排版。每個段落與重點之間必須有「空行」間距，讓操盤手能在一秒內看懂重點。絕不可輸出密密麻麻的長篇大論！";

                await LLMService.StreamChat(currentApiKey, new List<object> { new { role = "system", content = txtLivePrompt.Text + formatRule }, new { role = "user", content = sb.ToString() } }, (chunk) => {
                    this.Invoke(new Action(() => { txtLiveAiDiagnosis.AppendText(chunk); txtLiveAiDiagnosis.ScrollToCaret(); }));
                });
            }
            catch (Exception ex)
            {
                AppLogger.Log("BtnLiveAiScan_Click 失敗", ex);
                txtLiveAiDiagnosis.Text = $"掃描失敗: {ex.Message}";
            }
            finally
            {
                btnLiveAiScan.Enabled = true;
                btnLiveAiScan.Text = "⚡ AI 診斷即時動能 + VWAP + Fib";
            }
        }
        // ── 投資組合 ──────────────────────────────────────────────────────────
        private async void BtnRunPortfolio_Click(object sender, EventArgs e)
        {
            currentApiKey = txtApiKey.Text.Trim();
            if (string.IsNullOrEmpty(currentApiKey)) { MessageBox.Show("請輸入 API Key"); return; }
            string[] tickers = txtWatchlist.Text.Split(new[] { ',', ' ', '，' }, StringSplitOptions.RemoveEmptyEntries);
            if (tickers.Length < 2) { MessageBox.Show("請輸入至少兩檔股票。"); return; }
            btnRunPortfolio.Enabled = false; progressBar.Visible = true; txtPortfolioLog.Clear();
            try
            {
                var stats = new List<dynamic>(); var sb = new StringBuilder("【雷達近期技術狀態】\nTicker,Name,Close,RSI,MACD,Trend\n");
                foreach (var tk in tickers)
                {
                    string t = tk.Trim().ToUpper(); lblStatus.Text = $"取得 {t} 數據...";
                    var data = await YahooDataService.FetchYahoo(t, "1d", "6mo");
                    if (data.Count > 0)
                    {
                        IndicatorEngine.CalculateAll(data); var last = data.Last();
                        string name = await MarketInfoService.GetCompanyName(t, !t.EndsWith(".TW") && !t.EndsWith(".TWO"));
                        string trend = last.EMA_200 > 0 ? (last.EMA_50 > last.EMA_200 ? "多頭" : "空頭") : "-";
                        stats.Add(new { 代號 = t, 名稱 = string.IsNullOrEmpty(name) ? t : name, 收盤 = last.Close.ToString("F2"), RSI = last.RSI.ToString("F1"), MACD = last.MACD_Hist.ToString("F2"), 趨勢 = trend });
                        sb.AppendLine($"{t},{name},{last.Close:F2},{last.RSI:F1},{last.MACD_Hist:F2},{trend}");
                    }
                }
                dgvPortfolio.DataSource = stats;
                lblStatus.Text = "AI MPT 最佳化運算..."; txtPortfolioLog.Text = "AI 避險基金經理人計算相關性與權重配置...\r\n\r\n";
                await LLMService.StreamChat(currentApiKey, new List<object> { new { role = "system", content = txtPortfolioPrompt.Text }, new { role = "user", content = sb.ToString() } },
                    (chunk) => this.Invoke(new Action(() => { txtPortfolioLog.AppendText(chunk); txtPortfolioLog.ScrollToCaret(); })));
                lblStatus.Text = "投資組合最佳化完成。";
            }
            catch (OperationCanceledException)
            {
                txtPortfolioLog.Text = "⏹ 已取消。";
            }
            catch (Exception ex)
            {
                AppLogger.Log("BtnRunPortfolio_Click 失敗", ex);
                txtPortfolioLog.Text = $"失敗: {ex.Message}";
                ShowToast($"❌ 組合分析失敗：{ex.Message.Substring(0, Math.Min(ex.Message.Length, 50))}", Color.FromArgb(120, 30, 20));
            }
            finally { btnRunPortfolio.Enabled = true; progressBar.Visible = false; }
        }

        // ── 交易日誌 CRUD ─────────────────────────────────────────────────────
        private void LoadJournal()
        {
            try
            {
                journalEntries = ExcelJournalManager.LoadAll();
                AlertEngine.SyncFromJournal(journalEntries);
                RefreshJournalGrid();
                RefreshAlertGrid();
            }
            catch (Exception ex) { AppLogger.Log("LoadJournal 失敗", ex); }
        }

        private void RefreshJournalGrid()
        {
            dgvJournal.DataSource = journalEntries.Select(j => new
            {
                ID = j.Id,
                日期 = j.TradeDate.ToString("yyyy-MM-dd"),
                代號 = j.Ticker,
                方向 = j.Direction,
                進場 = j.EntryPrice.ToString("F2"),
                停損 = j.StopLossPrice > 0 ? j.StopLossPrice.ToString("F2") : "-",
                目標 = j.TargetPrice > 0 ? j.TargetPrice.ToString("F2") : "-",
                風報比 = j.RiskRewardRatio > 0 ? $"1:{j.RiskRewardRatio:F1}" : "-",
                出場 = j.ExitPrice > 0 ? j.ExitPrice.ToString("F2") : "-",
                股數 = j.Quantity.ToString("F0"),
                損益 = j.ExitPrice > 0 ? j.PnL.ToString("F0") : "-",
                報酬率 = j.ExitPrice > 0 ? j.ReturnPct.ToString("P2") : "-",
                狀態 = j.Status,
                AI建議 = j.AiSuggestion,
                決策 = j.MyDecision,
                備註 = j.Notes
            }).ToList();

            foreach (DataGridViewRow row in dgvJournal.Rows)
            {
                string status = row.Cells["狀態"].Value?.ToString() ?? "";
                if (status == "獲利") row.DefaultCellStyle.ForeColor = Color.SpringGreen;
                else if (status == "虧損") row.DefaultCellStyle.ForeColor = Color.LightCoral;
                else row.DefaultCellStyle.ForeColor = Color.Gold;
            }

            // 統計摘要
            var closed = journalEntries.Where(j => j.ExitPrice > 0).ToList();
            int wins = closed.Count(j => j.PnL > 0);
            double totalPnL = closed.Sum(j => j.PnL);
            int followAI = journalEntries.Count(j => j.MyDecision == "Follow");
            int againstAI = journalEntries.Count(j => j.MyDecision == "Against");
            lblJournalStats.Text = $"總交易: {journalEntries.Count} | 已平倉: {closed.Count} | 勝率: {(closed.Count > 0 ? (double)wins / closed.Count : 0):P1} | 總損益: {totalPnL:N0} | 跟AI: {followAI} vs 逆AI: {againstAI}";

            // 更新績效曲線圖
            var equityCtrl = contentPanels[4]?.Controls
                .OfType<Panel>()
                .SelectMany(p => p.Controls.OfType<EquityCurvePanel>())
                .FirstOrDefault();
            equityCtrl?.UpdateData(journalEntries);
        }

        private void BtnJournalAdd_Click(object sender, EventArgs e)
        {
            string aiSugg = historyData.Count > 0 ? historyData.Last().AgentAction : "-";
            // 預帶 ATR 停損建議
            RiskRewardSuggestion rrSugg = null;
            if (historyData.Count > 0) rrSugg = StopLossSuggestionEngine.CalcByATR(historyData, "Buy");
            using (var dlg = new JournalEditDialog(null, txtTicker.Text.Trim().ToUpper(), aiSugg, rrSugg))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    var entry = dlg.GetEntry();
                    entry.Id = ExcelJournalManager.NextId(journalEntries);
                    journalEntries.Add(entry);
                    ExcelJournalManager.SaveAll(journalEntries);
                    AlertEngine.SyncFromJournal(journalEntries);
                    RefreshJournalGrid();
                    RefreshAlertGrid();
                }
            }
        }

        private void BtnJournalEdit_Click(object sender, EventArgs e)
        {
            if (dgvJournal.SelectedRows.Count == 0) { MessageBox.Show("請先選取一筆記錄"); return; }
            int id = (int)dgvJournal.SelectedRows[0].Cells["ID"].Value;
            var entry = journalEntries.FirstOrDefault(j => j.Id == id);
            if (entry == null) return;
            using (var dlg = new JournalEditDialog(entry, entry.Ticker, entry.AiSuggestion, null))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    var updated = dlg.GetEntry(); updated.Id = id;
                    int idx = journalEntries.FindIndex(j => j.Id == id);
                    if (idx >= 0) journalEntries[idx] = updated;
                    ExcelJournalManager.SaveAll(journalEntries);
                    AlertEngine.SyncFromJournal(journalEntries);
                    RefreshJournalGrid();
                    RefreshAlertGrid();
                }
            }
        }

        private void BtnJournalDelete_Click(object sender, EventArgs e)
        {
            if (dgvJournal.SelectedRows.Count == 0) { MessageBox.Show("請先選取一筆記錄"); return; }
            if (MessageBox.Show("確定刪除此筆記錄？", "確認", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            int id = (int)dgvJournal.SelectedRows[0].Cells["ID"].Value;
            journalEntries.RemoveAll(j => j.Id == id);
            ExcelJournalManager.SaveAll(journalEntries);
            AlertEngine.SyncFromJournal(journalEntries);
            RefreshJournalGrid();
            RefreshAlertGrid();
        }

        // ── 股票篩選器 ────────────────────────────────────────────────────────
        private async void BtnRunScreener_Click(object sender, EventArgs e)
        {
            string raw = txtScreenerTickers.Text.Trim();
            if (string.IsNullOrEmpty(raw)) { MessageBox.Show("請輸入股票代號清單"); return; }
            var tickers = raw.Split(new[] { ',', ' ', '，', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim()).Distinct().ToList();
            if (tickers.Count == 0) return;

            var criteria = new ScreenerCriteria
            {
                RSI_Min = chkRsiFilter.Checked ? (double)numRsiMin.Value : 0,
                RSI_Max = chkRsiFilter.Checked ? (double)numRsiMax.Value : 100,
                MACD_Positive = chkMacdPos.Checked,
                MACD_CrossUp = chkMacdCross.Checked,
                Above_EMA50 = chkAboveEma50.Checked,
                Above_EMA200 = chkAboveEma200.Checked,
                BB_Breakout = chkBBBreakout.Checked,
                BB_Oversold = chkBBOversold.Checked,
                Volume_Min_Ratio = (double)numVolRatio.Value
            };

            btnRunScreener.Enabled = false; btnRunScreener.Text = "🔍 篩選中...";
            txtScreenerLog.Text = $"開始掃描 {tickers.Count} 檔股票...\r\n";
            dgvScreener.DataSource = null;
            try
            {
                var results = await ScreenerEngine.ScanAsync(tickers, criteria,
                    (msg) => this.Invoke(new Action(() => { txtScreenerLog.AppendText(msg + "\r\n"); txtScreenerLog.ScrollToCaret(); })));

                if (results.Count > 0)
                {
                    dgvScreener.DataSource = results.Select(r => new
                    {
                        代號 = r.Ticker,
                        名稱 = r.CompanyName,
                        收盤 = r.Close.ToString("F2"),
                        RSI = r.RSI.ToString("F1"),
                        MACD柱 = r.MACD_Hist.ToString("F3"),
                        趨勢 = r.Trend,
                        形態 = r.Pattern,
                        BBW = r.BB_Width.ToString("F1") + "%",
                        符合條件 = r.MatchScore,
                        條件明細 = string.Join(" | ", r.MatchedRules)
                    }).ToList();

                    // 高分列標色
                    foreach (DataGridViewRow row in dgvScreener.Rows)
                    {
                        int score = (int)row.Cells["符合條件"].Value;
                        if (score >= 4) row.DefaultCellStyle.ForeColor = Color.SpringGreen;
                        else if (score >= 2) row.DefaultCellStyle.ForeColor = Color.Gold;
                    }

                    txtScreenerLog.AppendText($"\r\n✅ 篩選完成！找到 {results.Count} 檔符合條件的股票。");
                }
                else
                {
                    txtScreenerLog.AppendText("\r\n⚠️ 沒有股票符合所有條件，請放寬篩選設定。");
                }
                lblStatus.Text = $"篩選完成，符合 {results.Count} 檔。";
            }
            catch (Exception ex) { AppLogger.Log("BtnRunScreener_Click 失敗", ex); txtScreenerLog.AppendText($"\r\n失敗: {ex.Message}"); }
            finally { btnRunScreener.Enabled = true; btnRunScreener.Text = "🔍 執行篩選"; }
        }

        // ── 停損停利計算 ──────────────────────────────────────────────────────────
        private void CalcRiskReward(string method)
        {
            if (historyData.Count < 15) { txtRiskRewardResult.Text = "⚠️ 請先載入歷史數據。"; return; }
            string direction = cbRRDirection.SelectedIndex == 0 ? "Buy" : "Sell";
            RiskRewardSuggestion s = method == "ATR"
                ? StopLossSuggestionEngine.CalcByATR(historyData, direction)
                : StopLossSuggestionEngine.CalcByFibonacci(historyData, direction);
            if (s == null) { txtRiskRewardResult.Text = "⚠️ 數據不足，無法計算。"; return; }
            txtRiskRewardResult.Text = s.Summary + $"\r\n風報比(T1): 1:{Math.Abs(s.Target1 - s.EntryPrice) / Math.Max(0.01, Math.Abs(s.EntryPrice - s.StopLoss)):F1}";
        }

        // ── 警報格表刷新 ──────────────────────────────────────────────────────────
        private void RefreshAlertGrid()
        {
            try
            {
                var alerts = AlertEngine.GetActiveAlerts();
                dgvAlerts.DataSource = journalEntries
                    .Where(j => j.ExitPrice == 0 && (j.StopLossPrice > 0 || j.TargetPrice > 0))
                    .Select(j =>
                    {
                        var a = alerts.FirstOrDefault(x => x.JournalId == j.Id);
                        return new
                        {
                            JournalId = j.Id,
                            代號 = j.Ticker,
                            方向 = j.Direction,
                            進場價 = j.EntryPrice.ToString("F2"),
                            停損價 = j.StopLossPrice > 0 ? j.StopLossPrice.ToString("F2") : "-",
                            目標價 = j.TargetPrice > 0 ? j.TargetPrice.ToString("F2") : "-",
                            狀態 = a != null ? (a.StopTriggered ? "🚨 停損觸發!" : a.TargetTriggered ? "🎯 目標達到!" : "🟢 監控中") : "⏸ 等待",
                            觸發時間 = a?.TriggeredAt?.ToString("HH:mm:ss") ?? "-"
                        };
                    }).ToList();

                // 觸發列紅色標記
                foreach (DataGridViewRow row in dgvAlerts.Rows)
                {
                    string status = row.Cells["狀態"].Value?.ToString() ?? "";
                    if (status.Contains("停損")) row.DefaultCellStyle.BackColor = Color.FromArgb(80, 0, 0);
                    else if (status.Contains("目標")) row.DefaultCellStyle.BackColor = Color.FromArgb(0, 60, 0);
                }
                lblAlertCount.Text = $"活躍: {alerts.Count} 筆";

                // ← 修正：有活躍警報時確保 timer 在執行（用戶新增警報後會自動啟動）
                if (alerts.Count > 0 && !alertTimer.Enabled)
                    alertTimer.Start();
                else if (alerts.Count == 0 && alertTimer.Enabled)
                    alertTimer.Stop();
            }
            catch (Exception ex) { AppLogger.Log("RefreshAlertGrid 失敗", ex); }
        }

        // ── 警報計時器 Tick ───────────────────────────────────────────────────────
        // 修正1：無活躍警報時自動暫停，不空跑掃描
        // 修正2：MessageBox 阻斷改為 Toast（不中斷 UI）
        private void AlertTimer_Tick(object sender, EventArgs e)
        {
            var alerts = AlertEngine.GetActiveAlerts();
            if (alerts.Count == 0)
            {
                alertTimer.Stop();  // ← 沒有活躍警報時自動停止，節省 CPU
                return;
            }

            foreach (var a in alerts.ToList())
            {
                double price = 0;
                if (currentLiveData.Count > 0 && _currentLiveTicker == a.Ticker)
                    price = currentLiveData.Last().Close;
                else if (historyData.Count > 0 && txtTicker.Text.Trim().ToUpper() == a.Ticker)
                    price = historyData.Last().Close;
                if (price <= 0) continue;

                var triggered = AlertEngine.CheckPrice(a.Ticker, price);
                foreach (var t in triggered)
                {
                    string emoji = t.TriggeredType == "StopLoss" ? "🚨 停損觸發" : "🎯 目標達到";
                    string toastMsg = $"{emoji}  {t.Ticker} 現價 {price:F2}  " +
                                     (t.TriggeredType == "StopLoss"
                                         ? $"停損位: {t.StopLossPrice:F2}"
                                         : $"目標位: {t.TargetPrice:F2}");
                    Color toastColor = t.TriggeredType == "StopLoss"
                        ? Color.FromArgb(150, 40, 25)
                        : Color.FromArgb(20, 110, 50);
                    ShowToast(toastMsg, toastColor);
                }
                if (triggered.Count > 0) RefreshAlertGrid();
            }
        }

        // ── 即時警報價格檢查（從 RefreshLive 呼叫） ────────────────────────────
        private void CheckAlertsForPrice(string ticker, double price)
        {
            var triggered = AlertEngine.CheckPrice(ticker, price);
            if (triggered.Count == 0) return;
            foreach (var a in triggered)
            {
                string emoji = a.TriggeredType == "StopLoss" ? "🚨 停損觸發" : "🎯 目標達到";
                string msg = $"{emoji}  {ticker} 現價 {price:F2}  " +
                             (a.TriggeredType == "StopLoss"
                                 ? $"| 停損位: {a.StopLossPrice:F2}"
                                 : $"| 目標位: {a.TargetPrice:F2}");
                Color toastColor = a.TriggeredType == "StopLoss"
                    ? Color.FromArgb(150, 40, 25)
                    : Color.FromArgb(20, 110, 50);
                ShowToast(msg, toastColor);
            }
            if (_currentPage == 4) RefreshAlertGrid();
        }

        // ── 市場情緒刷新 ─────────────────────────────────────────────────────────
        private async Task RefreshSentimentAsync(bool force = false)
        {
            if (btnRefreshSentiment == null) return;
            btnRefreshSentiment.Enabled = false; btnRefreshSentiment.Text = "🔄 取得中...";
            try
            {
                var s = await MarketSentimentService.FetchAsync(force);
                _lastSentiment = s;
                lblVIX.Text = $"{s.VIX:F2}";
                lblVIX.ForeColor = s.VIX > 35 ? Color.OrangeRed : s.VIX > 25 ? Color.Orange : s.VIX > 15 ? Color.Gold : Color.SpringGreen;
                lblVIXChange.Text = $"{(s.VIXChange >= 0 ? "+" : "")}{s.VIXChange:F2}  {s.VIXLevel}";
                lblVIXChange.ForeColor = s.VIXChange > 0 ? Color.OrangeRed : Color.SpringGreen;
                lblFGScore.Text = $"{s.FearGreedScore:F0} / 100";
                lblFGScore.ForeColor = s.FearGreedScore >= 70 ? Color.OrangeRed :
                                       s.FearGreedScore >= 50 ? Color.Gold :
                                       s.FearGreedScore >= 30 ? Color.LightSkyBlue : Color.MediumOrchid;
                lblFGLabel.Text = s.FearGreedLabel;
                lblSPChange.Text = $"S&P 500  {(s.SP500ChangePct >= 0 ? "+" : "")}{s.SP500ChangePct:F2}%";
                lblSPChange.ForeColor = s.SP500ChangePct >= 0 ? Color.SpringGreen : Color.LightCoral;
                lblNdqChange.Text = $"Nasdaq  {(s.NasdaqChangePct >= 0 ? "+" : "")}{s.NasdaqChangePct:F2}%";
                lblNdqChange.ForeColor = s.NasdaqChangePct >= 0 ? Color.SpringGreen : Color.LightCoral;
                lblPosAdvice.Text = "💡 倉位建議: " + s.AiPositionAdvice;
                txtSentimentDetail.Text = s.Summary + "\r\n\r\n" +
                    $"=== VIX 解讀 ===\r\n" +
                    $"VIX < 15: 市場極度樂觀，歷史上常是短期頂部附近\r\n" +
                    $"VIX 15~25: 正常波動，可正常操作\r\n" +
                    $"VIX 25~35: 市場恐慌，宜降低倉位，嚴守停損\r\n" +
                    $"VIX > 35: 極度恐慌，歷史上常是中長期買點\r\n\r\n" +
                    $"=== Fear & Greed 合成指數說明 ===\r\n" +
                    $"權重組成: VIX反轉(35%) + S&P動能(25%) + Nasdaq動能(20%) + VIX動量(10%) + 市場廣度(10%)\r\n" +
                    $"本指數每30分鐘自動快取，避免頻繁請求。\r\n";
                pnlFearGreed?.Invalidate();
                lblStatus.Text = $"市場情緒更新完成  VIX={s.VIX:F2}  F&G={s.FearGreedScore:F0}/100";
            }
            catch (Exception ex) { AppLogger.Log("RefreshSentimentAsync 失敗", ex); }
            finally
            {
                if (btnRefreshSentiment != null)
                { btnRefreshSentiment.Enabled = true; btnRefreshSentiment.Text = "🔄 立即刷新"; }
            }
        }

        // ── 提示詞 ────────────────────────────────────────────────────────────
        private void LoadPrompts()
        {
            string defHist = "📊 STOCK ANALYSIS STACK v3.0\n【角色】50年經驗長期投資者，數據導向，結合技術面+基本面+多時間框架共振。\n【任務】綜合提供的歷史技術數據、多時間框架共振信號、新聞情緒與基本面，產生研究級投資分析。\n【規則】①不編造數字 ②同時提供多空情境 ③特別重視多時間框架共振分數 ④若P/E或Yield為0請忽略\n🚨必須輸出合法JSON: { \"debate_log\": \"...\", \"results\": [ { \"Date\": \"MMdd\", \"Action\": \"Buy/Sell/Hold\", \"Reasoning\": \"...\" } ] }";
            string defLive = "你是極短線當沖專家。分析即時Tick走勢、新聞與Fibonacci回調位，判斷「真突破」、「誘多/誘空」或「盤整」，特別注意VWAP和Fib支撐壓制，給出秒級操作建議。🚨用繁體中文回答。";
            string defPort = "你是避險基金經理人，精通MPT。根據多檔股票近期技術指標，評估相關性與風險，給出具體資金分配權重建議(如台積電40%,NVDA30%)，說明如何降低非系統性風險。🚨用繁體中文回答。";
            txtSystemPrompt.Text = File.Exists(histPromptFile) ? File.ReadAllText(histPromptFile) : defHist;
            txtLivePrompt.Text = File.Exists(livePromptFile) ? File.ReadAllText(livePromptFile) : defLive;
            txtPortfolioPrompt.Text = File.Exists(portfolioPromptFile) ? File.ReadAllText(portfolioPromptFile) : defPort;
        }

        private void BtnSavePrompts_Click(object sender, EventArgs e)
        {
            try
            {
                File.WriteAllText(histPromptFile, txtSystemPrompt.Text);
                File.WriteAllText(livePromptFile, txtLivePrompt.Text);
                File.WriteAllText(portfolioPromptFile, txtPortfolioPrompt.Text);
                MessageBox.Show("提示詞已儲存！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { AppLogger.Log("BtnSavePrompts_Click 失敗", ex); MessageBox.Show($"儲存失敗：{ex.Message}", "錯誤"); }
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  ETF 資訊卡
        // ═══════════════════════════════════════════════════════════════════════

        /// <summary>建立 ETF 資訊卡 Panel（初始隱藏，偵測到 ETF 才顯示）</summary>
        private Panel BuildEtfCardPanel()
        {
            var pnl = new Panel
            {
                Dock = DockStyle.Top,
                Height = 0,     // 初始折疊
                Visible = false,
                BackColor = Color.FromArgb(20, 40, 65),
                Padding = new Padding(14, 8, 14, 8)
            };

            // 第一行：名稱 + 類別 Badge
            var pnlLine1 = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 28,
                BackColor = Color.Transparent,
                WrapContents = false
            };
            lblEtfBadge = new Label
            {
                Text = "📦 ETF",
                AutoSize = false,
                Width = 55,
                Height = 22,
                BackColor = Color.FromArgb(0, 100, 180),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(0, 2, 6, 0)
            };
            var lblEtfName = new Label
            {
                Name = "lblEtfName",
                Text = "",
                AutoSize = true,
                ForeColor = Color.White,
                Font = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold),
                Margin = new Padding(0, 2, 0, 0)
            };
            var lblEtfCat = new Label
            {
                Name = "lblEtfCat",
                Text = "",
                AutoSize = true,
                ForeColor = Color.LightSkyBlue,
                Font = new Font("Segoe UI", 9.5F),
                Margin = new Padding(10, 4, 0, 0)
            };
            pnlLine1.Controls.AddRange(new Control[] { lblEtfBadge, lblEtfName, lblEtfCat });

            // 第二行：AUM / 費率 / 殖利率 / 今年報酬
            var pnlLine2 = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 24,
                BackColor = Color.Transparent,
                WrapContents = false
            };
            var lblEtfStats = new Label
            {
                Name = "lblEtfStats",
                Text = "",
                AutoSize = true,
                ForeColor = Color.LightGoldenrodYellow,
                Font = new Font("Consolas", 10F)
            };
            pnlLine2.Controls.Add(lblEtfStats);

            // 第三行：前三大持股
            var lblEtfHoldings = new Label
            {
                Name = "lblEtfHoldings",
                Text = "",
                Dock = DockStyle.Top,
                Height = 20,
                ForeColor = Color.LightGray,
                Font = new Font("Segoe UI", 9F)
            };

            pnl.Controls.Add(lblEtfHoldings);
            pnl.Controls.Add(pnlLine2);
            pnl.Controls.Add(pnlLine1);
            return pnl;
        }

        /// <summary>用最新 EtfInfo 更新 ETF 資訊卡</summary>
        private void UpdateEtfCard(EtfInfo info)
        {
            if (pnlEtfCard == null || info == null) return;
            if (InvokeRequired) { Invoke(new Action(() => UpdateEtfCard(info))); return; }

            // 填資料
            var lblName = pnlEtfCard.Controls.Find("lblEtfName", true).FirstOrDefault() as Label;
            var lblCat = pnlEtfCard.Controls.Find("lblEtfCat", true).FirstOrDefault() as Label;
            var lblStats = pnlEtfCard.Controls.Find("lblEtfStats", true).FirstOrDefault() as Label;
            var lblHoldings = pnlEtfCard.Controls.Find("lblEtfHoldings", true).FirstOrDefault() as Label;

            if (lblName != null) lblName.Text = string.IsNullOrEmpty(info.Name) ? info.Ticker : info.Name;
            if (lblCat != null) lblCat.Text = string.IsNullOrEmpty(info.Category) ? "" : $"  [{info.Category}]";

            // AUM 格式化
            string aum = info.TotalAssets >= 1e9 ? $"AUM ${info.TotalAssets / 1e9:F1}B"
                       : info.TotalAssets >= 1e6 ? $"AUM ${info.TotalAssets / 1e6:F0}M"
                       : info.TotalAssets > 0 ? $"AUM ${info.TotalAssets:N0}"
                       : "AUM N/A";
            string exp = info.ExpenseRatio > 0 ? $"費率 {info.ExpenseRatio:P2}" : "費率 N/A";
            string yld = info.DividendYield > 0 ? $"殖利率 {info.DividendYield:P2}" : "";
            string ytd = info.YtdReturn != 0 ? $"YTD {info.YtdReturn:+0.0%;-0.0%}" : "";
            if (lblStats != null) lblStats.Text = string.Join("   ", new[] { aum, exp, yld, ytd }.Where(x => !string.IsNullOrEmpty(x)));

            if (lblHoldings != null && info.TopHoldings.Count > 0)
                lblHoldings.Text = "前三大: " + string.Join("  ", info.TopHoldings.Take(3).Select(h => $"{h.Name} {h.Pct:P0}"));
            else if (lblHoldings != null) lblHoldings.Text = "";

            // 展開卡片
            pnlEtfCard.Height = info.TopHoldings.Count > 0 ? 78 : 58;
            pnlEtfCard.Visible = true;
            if (lblEtfBadge != null) lblEtfBadge.Visible = true;
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  標的比較（疊加報酬率曲線）
        // ═══════════════════════════════════════════════════════════════════════
        private async Task RunComparisonAsync(string t1, string t2, string range)
        {
            if (string.IsNullOrEmpty(t1) || string.IsNullOrEmpty(t2))
            {
                ShowToast("⚠️ 請在比較欄位輸入第二個標的代號", Color.FromArgb(100, 70, 20));
                return;
            }
            btnRunCompare.Enabled = false;
            btnRunCompare.Text = "⏳ 載入...";
            lblStatus.Text = $"比較 {t1} vs {t2} ({range})...";
            try
            {
                var data = await ComparisonService.FetchComparisonAsync(t1, t2, range);
                if (data.Count < 2)
                {
                    ShowToast($"⚠️ {t1} 或 {t2} 資料不足，無法比較", Color.FromArgb(100, 70, 20));
                    return;
                }

                // 把比較數據疊加到 historyChart（建立假的 MarketData 列表用於顯示）
                // 在 txtDebateLog 輸出統計摘要
                double r1Final = data.Last().R1, r2Final = data.Last().R2;
                double maxR1 = data.Max(d => d.R1), maxR2 = data.Max(d => d.R2);
                double minR1 = data.Min(d => d.R1), minR2 = data.Min(d => d.R2);

                string winner = r1Final >= r2Final ? t1 : t2;
                double diff = Math.Abs(r1Final - r2Final);

                var sb = new StringBuilder();
                sb.AppendLine($"╔══ 標的比較：{t1} vs {t2}  [{range}] ══╗");
                sb.AppendLine($"  {t1,-8}  累計: {r1Final:+0.00%;-0.00%}  最高: {maxR1:+0.0%;-0.0%}  最低: {minR1:+0.0%;-0.0%}");
                sb.AppendLine($"  {t2,-8}  累計: {r2Final:+0.00%;-0.00%}  最高: {maxR2:+0.0%;-0.0%}  最低: {minR2:+0.0%;-0.0%}");
                sb.AppendLine($"  勝者: {winner}  差距: {diff:0.00%}");
                sb.AppendLine();
                sb.AppendLine("日期          " + t1.PadRight(10) + t2.PadRight(10));
                sb.AppendLine("──────────────────────────────");
                // 每月取樣
                DateTime lastPrint = DateTime.MinValue;
                foreach (var (dt, r1, r2) in data)
                {
                    if ((dt - lastPrint).TotalDays >= 28)
                    {
                        sb.AppendLine($"{dt:yyyy-MM-dd}  {r1,+8:0.00%}  {r2,+8:0.00%}");
                        lastPrint = dt;
                    }
                }
                sb.AppendLine($"{data.Last().Date:yyyy-MM-dd}  {r1Final,+8:0.00%}  {r2Final,+8:0.00%}  ← 最新");
                sb.AppendLine($"╚════════════════════════════╝");

                if (txtDebateLog != null)
                    txtDebateLog.Text = sb.ToString();

                lblStatus.Text = $"✅ 比較完成：{t1} {r1Final:+0.0%;-0.0%} vs {t2} {r2Final:+0.0%;-0.0%}";
                ShowToast($"📊 {winner} 表現較佳  差距 {diff:0.0%}", Color.FromArgb(20, 70, 120));
            }
            catch (Exception ex)
            {
                AppLogger.Log("RunComparisonAsync 失敗", ex);
                ShowToast($"❌ 比較失敗：{ex.Message}", Color.FromArgb(120, 30, 20));
            }
            finally
            {
                btnRunCompare.Enabled = true;
                btnRunCompare.Text = "▶ 比較走勢";
            }
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  板塊輪動面板
        // ═══════════════════════════════════════════════════════════════════════

        /// <summary>建立板塊輪動 Panel（放在 Tab7 右側下方）</summary>
        private Panel BuildSectorRotationPanel()
        {
            var pnl = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 300,
                BackColor = Color.FromArgb(16, 20, 30),
                Padding = new Padding(12, 8, 12, 8)
            };

            var lblTitle = new Label
            {
                Text = "🏭 板塊輪動快照（美股 11 個板塊 ETF）",
                Dock = DockStyle.Top,
                Height = 28,
                ForeColor = Color.Gold,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold)
            };

            var btnRefSector = new Button
            {
                Text = "🔄 刷新板塊",
                Dock = DockStyle.Top,
                Height = 34,
                BackColor = Color.FromArgb(0, 100, 100),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnRefSector.FlatAppearance.BorderSize = 0;

            pnlSectorRotation = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };
            pnlSectorRotation.Paint += DrawSectorBars;

            btnRefSector.Click += async (s, e) =>
            {
                btnRefSector.Enabled = false; btnRefSector.Text = "⏳ 載入中...";
                try
                {
                    _sectorData = await SectorRotationService.FetchSnapshotAsync();
                    pnlSectorRotation.Invalidate();
                    lblStatus.Text = $"板塊快照更新完成 ({DateTime.Now:HH:mm})";
                }
                catch (Exception ex) { AppLogger.Log("板塊刷新失敗", ex); }
                finally { btnRefSector.Enabled = true; btnRefSector.Text = "🔄 刷新板塊"; }
            };

            pnl.Controls.Add(pnlSectorRotation);
            pnl.Controls.Add(btnRefSector);
            pnl.Controls.Add(lblTitle);
            return pnl;
        }

        // 儲存最後一次板塊資料（繪製用）
        private List<(string Ticker, string Name, string Emoji, double ChangePct, double Price)> _sectorData = null;

        private void DrawSectorBars(object sender, PaintEventArgs e)
        {
            if (_sectorData == null || _sectorData.Count == 0)
            {
                using var font = new Font("Segoe UI", 10F);
                e.Graphics.DrawString("點擊「刷新板塊」載入數據", font, Brushes.Gray, 10, 10);
                return;
            }

            var g = e.Graphics;
            var panel = sender as Panel;
            int w = panel.Width, h = panel.Height;
            int rowH = Math.Max(16, (h - 4) / _sectorData.Count);
            double maxAbs = Math.Max(_sectorData.Max(x => Math.Abs(x.ChangePct)), 0.01);

            using var nameFont = new Font("Segoe UI", 9.5F);
            using var pctFont = new Font("Consolas", 9.5F, FontStyle.Bold);
            using var grayBrush = new SolidBrush(Color.FromArgb(60, 60, 70));

            for (int i = 0; i < _sectorData.Count; i++)
            {
                var (ticker, name, emoji, chg, price) = _sectorData[i];
                int y = i * rowH + 2;

                // 背景條
                g.FillRectangle(grayBrush, 0, y, w, rowH - 1);

                // 顏色條（正漲 = 綠，下跌 = 紅）
                double pct = Math.Abs(chg) / maxAbs;
                int barW = (int)((w - 160) * pct);
                int barX = chg >= 0 ? 140 : 140 + (int)((w - 160) * (Math.Abs(chg) / maxAbs));
                Color barColor = chg >= 0
                    ? Color.FromArgb(40, Math.Min(255, 60 + (int)(195 * pct)), 60)
                    : Color.FromArgb(Math.Min(255, 60 + (int)(195 * pct)), 40, 40);

                using var barBrush = new SolidBrush(barColor);
                int finalBarX = chg >= 0 ? 140 : 140;
                int finalBarW = (int)((w - 160) * Math.Abs(chg) / maxAbs);
                if (chg < 0) finalBarX = 140 + (w - 160) - finalBarW;
                if (finalBarW > 0) g.FillRectangle(barBrush, finalBarX, y + 2, finalBarW, rowH - 5);

                // 名稱
                string label = $"{emoji} {name}";
                g.DrawString(label, nameFont, Brushes.LightGray, 4, y + 2);

                // 漲跌幅
                string pctStr = $"{chg:+0.00%;-0.00%}";
                Color pctColor = chg >= 0 ? Color.SpringGreen : Color.LightCoral;
                using var pctBrush = new SolidBrush(pctColor);
                var pctSize = g.MeasureString(pctStr, pctFont);
                g.DrawString(pctStr, pctFont, pctBrush, w - (int)pctSize.Width - 4, y + 2);
            }
        }

        private void BtnOpenChat_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtApiKey.Text.Trim())) { MessageBox.Show("請先輸入 API Key"); return; }
            if (chatFormInstance == null || chatFormInstance.IsDisposed)
            {
                Func<string> getCtx = () =>
                {
                    if (_currentPage == 0 && dgvHistory.SelectedRows.Count > 0 && historyData.Count > 0)
                    {
                        var d = historyData[dgvHistory.SelectedRows[0].Index];
                        return $"【歷史】{d.Date:yyyy-MM-dd} 收={d.Close} RSI={d.RSI:F1} 建議={d.AgentAction}";
                    }
                    if (_currentPage == 1 && currentLiveData.Count > 0)
                    {
                        var d = currentLiveData.Last();
                        return $"【即時】{d.Date:HH:mm} 收={d.Close} VWAP={d.VWAP:F2} 支撐={d.SupportLevel:F2}";
                    }
                    return "無特定上下文。";
                };
                chatFormInstance = new FloatingChatForm(txtApiKey.Text.Trim(), txtTicker.Text.Trim(), getCtx);
                chatFormInstance.Show(this);
            }
            else chatFormInstance.Focus();
        }
    }

    // ════════════════════════════════════════════════════════════════════════════
    //  交易日誌新增/編輯對話框
    // ════════════════════════════════════════════════════════════════════════════
    // JournalEditDialog 類別：請用此區塊替換原檔中同名類別的內容
    public class JournalEditDialog : Form
    {
        private DateTimePicker dtpDate;
        private TextBox txtTicker, txtNotes;
        private ComboBox cbDirection, cbDecision;
        private NumericUpDown numEntry, numExit, numQty, numStopLoss, numTarget;
        private Label lblAiSugg, lblRR;
        private readonly string _aiSuggestion;
        private readonly TradeJournalEntry _existing;

        // 新增自動計算相關控制項
        private ComboBox cbSizeMethod;
        private NumericUpDown numCapitalForCalc;
        private NumericUpDown numParam; // shrink% 或 percent%
        private NumericUpDown numLotSize;
        private Button btnAutoQty;

        public JournalEditDialog(TradeJournalEntry existing, string ticker, string aiSuggestion, RiskRewardSuggestion rrSugg)
        {
            _existing = existing; _aiSuggestion = aiSuggestion;
            Text = existing == null ? "新增交易記錄" : "修改/平倉";
            Size = new Size(460, 640);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.FromArgb(30, 32, 40); ForeColor = Color.White;
            Font = new Font("Segoe UI", 10F);

            var tlp = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 14,
                Padding = new Padding(15),
                BackColor = Color.Transparent
            };
            for (int i = 0; i < 14; i++) tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            dtpDate = new DateTimePicker
            {
                Format = DateTimePickerFormat.Short,
                BackColor = Color.FromArgb(45, 48, 55),
                ForeColor = Color.White,
                Dock = DockStyle.Fill
            };
            txtTicker = new TextBox
            {
                BackColor = Color.FromArgb(45, 48, 55),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Fill
            };
            cbDirection = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(45, 48, 55),
                ForeColor = Color.White,
                Dock = DockStyle.Fill
            };
            cbDirection.Items.AddRange(new object[] { "Buy", "Sell" }); cbDirection.SelectedIndex = 0;
            numEntry = MakeNum(0, 999999, 2); numExit = MakeNum(0, 999999, 2); numQty = MakeNum(0, 9999999, 0);
            numStopLoss = MakeNum(0, 999999, 2); numTarget = MakeNum(0, 999999, 2);
            cbDecision = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(45, 48, 55),
                ForeColor = Color.White,
                Dock = DockStyle.Fill
            };
            cbDecision.Items.AddRange(new object[] { "Follow", "Against", "N/A" }); cbDecision.SelectedIndex = 2;
            lblAiSugg = new Label
            {
                Text = aiSuggestion,
                ForeColor = Color.Gold,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            lblRR = new Label
            {
                Text = "-",
                ForeColor = Color.LightGreen,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Consolas", 9.5F)
            };
            txtNotes = new TextBox
            {
                BackColor = Color.FromArgb(45, 48, 55),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Fill
            };

            // 更新風報比顯示
            void UpdateRR()
            {
                decimal entry = numEntry.Value, stop = numStopLoss.Value, target = numTarget.Value;
                if (entry > 0 && stop > 0 && target > 0 && entry != stop)
                {
                    double rr = (double)Math.Abs(target - entry) / (double)Math.Abs(entry - stop);
                    double risk = (double)Math.Abs(entry - stop) / (double)entry;
                    lblRR.Text = $"風報比 1:{rr:F1}  停損幅度 {risk:P1}";
                    lblRR.ForeColor = rr >= 2 ? Color.SpringGreen : rr >= 1 ? Color.Gold : Color.LightCoral;
                }
                else lblRR.Text = "-";
            }
            numEntry.ValueChanged += (s, e) => UpdateRR();
            numStopLoss.ValueChanged += (s, e) => UpdateRR();
            numTarget.ValueChanged += (s, e) => UpdateRR();

            void AddRow(int r, string label, Control ctrl)
            {
                tlp.Controls.Add(new Label
                {
                    Text = label,
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleRight,
                    ForeColor = Color.LightGray
                }, 0, r);
                tlp.Controls.Add(ctrl, 1, r);
            }
            AddRow(0, "日期", dtpDate);
            AddRow(1, "代號", txtTicker);
            AddRow(2, "方向", cbDirection);
            AddRow(3, "進場價", numEntry);
            AddRow(4, "🛑 停損價", numStopLoss);
            AddRow(5, "🎯 目標價", numTarget);
            AddRow(6, "風報比", lblRR);
            AddRow(7, "出場價", numExit);
            AddRow(8, "股數", numQty);
            AddRow(9, "AI 建議", lblAiSugg);

            // 新增：自動計算列（row 10）
            // 內含：方法選擇（Kelly / 固定%）、資金、參數（shrink% 或 percent%）、lotSize、按鈕
            var pnlAuto = new FlowLayoutPanel { Dock = DockStyle.Fill, BackColor = Color.Transparent, WrapContents = false, AutoSize = false };
            cbSizeMethod = new ComboBox { Width = 130, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            cbSizeMethod.Items.AddRange(new object[] { "Kelly(半Kelly)", "固定% (Percent)" });
            cbSizeMethod.SelectedIndex = 0;
            numCapitalForCalc = new NumericUpDown { Minimum = 1000, Maximum = 1000000000, Value = 100000, Width = 120, DecimalPlaces = 0, Increment = 1000, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White };
            numParam = new NumericUpDown { Minimum = 0, Maximum = 100, Value = 50, Width = 70, DecimalPlaces = 1, Increment = 0.5M, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White }; // shrink% or percent%
            numLotSize = new NumericUpDown { Minimum = 1, Maximum = 1000000, Value = 1, Width = 70, DecimalPlaces = 0, BackColor = Color.FromArgb(45, 48, 55), ForeColor = Color.White };

            btnAutoQty = new Button { Text = "自動計算股數", Width = 120, Height = 28, BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnAutoQty.FlatAppearance.BorderSize = 0;
            btnAutoQty.Click += (s, e) =>
            {
                try
                {
                    double entry = (double)numEntry.Value;
                    double stop = (double)numStopLoss.Value;
                    if (entry <= 0 || stop <= 0) { MessageBox.Show("請先輸入進場價與停損價"); return; }
                    double capital = (double)numCapitalForCalc.Value;
                    int lot = (int)numLotSize.Value;
                    long qty = 0;
                    if (cbSizeMethod.SelectedIndex == 0)
                    {
                        // Kelly
                        var journals = ExcelJournalManager.LoadAll() ?? new List<TradeJournalEntry>();
                        double kelly = PositionSizingService.CalcKellyFromJournalEntries(journals);
                        double shrink = (double)numParam.Value / 100.0; // user enters 50 -> 0.5
                        if (kelly <= 0)
                        {
                            // fallback to fixed 2%
                            double frac = 0.02;
                            qty = PositionSizingService.CalcQuantityFromFraction(capital, frac, entry, stop, lot);
                        }
                        else
                        {
                            double frac = PositionSizingService.CalcPositionFractionKelly(kelly, shrink);
                            qty = PositionSizingService.CalcQuantityFromFraction(capital, frac, entry, stop, lot);
                        }
                    }
                    else
                    {
                        // 固定%
                        double percent = (double)numParam.Value / 100.0;
                        qty = PositionSizingService.CalcQuantityFixedPercent(capital, percent, entry, stop, lot);
                    }

                    if (qty <= 0)
                    {
                        MessageBox.Show("計算結果為 0 股（可能資金、停損或 lotSize 不合理）", "計算結果", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        numQty.Value = Math.Min(numQty.Maximum, qty);
                        MessageBox.Show($"建議股數：{qty}", "計算完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"自動計算失敗：{ex.Message}");
                }
            };

            // UI 排列：顯示說明文字
            var lblCap = new Label { Text = "資金", ForeColor = Color.Gray, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft };
            var lblParam = new Label { Text = "Shrink%/Percent%", ForeColor = Color.Gray, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft };
            var lblLot = new Label { Text = "Lot", ForeColor = Color.Gray, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft };

            pnlAuto.Controls.Add(cbSizeMethod);
            pnlAuto.Controls.Add(lblCap); pnlAuto.Controls.Add(numCapitalForCalc);
            pnlAuto.Controls.Add(lblParam); pnlAuto.Controls.Add(numParam);
            pnlAuto.Controls.Add(lblLot); pnlAuto.Controls.Add(numLotSize);
            pnlAuto.Controls.Add(btnAutoQty);

            AddRow(10, "自動計算", pnlAuto);

            AddRow(11, "我的決策", cbDecision);
            AddRow(12, "備註", txtNotes);

            var pnlBtn = new Panel { Dock = DockStyle.Bottom, Height = 55, BackColor = Color.FromArgb(22, 24, 32) };
            var btnOk = new Button
            {
                Text = "✅ 確定",
                Dock = DockStyle.Right,
                Width = 100,
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.OK
            };
            btnOk.FlatAppearance.BorderSize = 0;
            var btnCancel = new Button
            {
                Text = "取消",
                Dock = DockStyle.Right,
                Width = 80,
                BackColor = Color.FromArgb(80, 80, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            pnlBtn.Controls.Add(btnOk); pnlBtn.Controls.Add(btnCancel);

            Controls.Add(tlp); Controls.Add(pnlBtn);
            AcceptButton = btnOk; CancelButton = btnCancel;

            if (existing != null)
            {
                dtpDate.Value = existing.TradeDate; txtTicker.Text = existing.Ticker;
                if (cbDirection.Items.Contains(existing.Direction)) cbDirection.SelectedItem = existing.Direction;
                numEntry.Value = (decimal)existing.EntryPrice;
                if (existing.StopLossPrice > 0) numStopLoss.Value = (decimal)existing.StopLossPrice;
                if (existing.TargetPrice > 0) numTarget.Value = (decimal)existing.TargetPrice;
                if (existing.ExitPrice > 0) numExit.Value = (decimal)existing.ExitPrice;
                numQty.Value = (decimal)existing.Quantity;
                if (cbDecision.Items.Contains(existing.MyDecision)) cbDecision.SelectedItem = existing.MyDecision;
                txtNotes.Text = existing.Notes;
            }
            else
            {
                txtTicker.Text = ticker;
                // 預填 ATR 建議
                if (rrSugg != null)
                {
                    numEntry.Value = (decimal)rrSugg.EntryPrice;
                    numStopLoss.Value = (decimal)rrSugg.StopLoss;
                    numTarget.Value = (decimal)rrSugg.Target2;  // 預設 T2 (1:2)
                }
            }
            UpdateRR();
        }

        private NumericUpDown MakeNum(decimal min, decimal max, int decimals) =>
            new NumericUpDown
            {
                Minimum = min,
                Maximum = max,
                DecimalPlaces = decimals,
                Increment = decimals > 0 ? 0.1M : 1M,
                BackColor = Color.FromArgb(45, 48, 55),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Fill
            };

        public TradeJournalEntry GetEntry() => new TradeJournalEntry
        {
            TradeDate = dtpDate.Value.Date,
            Ticker = txtTicker.Text.Trim().ToUpper(),
            Direction = cbDirection.SelectedItem?.ToString() ?? "Buy",
            EntryPrice = (double)numEntry.Value,
            StopLossPrice = (double)numStopLoss.Value,
            TargetPrice = (double)numTarget.Value,
            ExitPrice = (double)numExit.Value,
            Quantity = (double)numQty.Value,
            AiSuggestion = _aiSuggestion,
            MyDecision = cbDecision.SelectedItem?.ToString() ?? "N/A",
            Notes = txtNotes.Text
        };
    }
}