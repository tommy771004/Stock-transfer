using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LLMAgentTrader
{
    // ════════════════════════════════════════════════════════════════
    //  台股當沖 AI 模擬交易引擎  ─  修正版
    //  修正說明：
    //  1. [CRITICAL] 競態條件：資金/庫存判斷改在 lock 內執行
    //  2. [CRITICAL] Fire-and-Forget：改用 SemaphoreSlim(1) 限制並發
    //  3. [HIGH]     股價加入 ±10% 漲跌停 + 最低價 0.01 保護
    //  4. [HIGH]     加入台股交易時段判斷（09:00 ~ 13:30）
    //  5. [HIGH]     加入 13:25 強制平倉機制（當沖核心規則）
    //  6. [HIGH]     加入手續費 0.1425% + 當沖證交稅 0.15%
    //  7. [MEDIUM]   庫存統一改為「張數(lots)」，底層不再存股數
    //  8. [LOW]      加入滑價模擬（0 ~ 0.05% 隨機）
    // ════════════════════════════════════════════════════════════════

    public class AIAutoTradeForm : Form
    {
        private TextBox txtSymbol;
        private NumericUpDown numDailyCapital;
        private NumericUpDown numTradeLots;
        private Button btnStart;
        private Button btnStop;
        private RichTextBox rtbLog;
        private Label lblStatus;
        private Label lblPosition;

        private AIAutoTradeEngine _engine;

        public AIAutoTradeForm()
        {
            InitializeComponent();
            _engine = new AIAutoTradeEngine();
            _engine.OnLogMessage += Engine_OnLogMessage;
            _engine.OnStatusUpdate += Engine_OnStatusUpdate;
        }

        private void InitializeComponent()
        {
            this.Text = "AI 當沖模擬交易（修正版）";
            this.Size = new Size(1100, 650);
            this.StartPosition = FormStartPosition.CenterScreen;

            Panel panelTop = new Panel { Dock = DockStyle.Top, Height = 130, Padding = new Padding(10) };

            panelTop.Controls.Add(new Label { Text = "股票代號:", Location = new Point(20, 20), AutoSize = true });
            txtSymbol = new TextBox { Location = new Point(100, 18), Width = 100, Text = "2330" };
            panelTop.Controls.Add(txtSymbol);

            panelTop.Controls.Add(new Label { Text = "每日本錢(元):", Location = new Point(220, 20), AutoSize = true });
            numDailyCapital = new NumericUpDown
            {
                Location = new Point(320, 18),
                Width = 120,
                Maximum = 100_000_000m,
                Minimum = 10_000m,
                Value = 250_000m,
                Increment = 10_000m
            };
            panelTop.Controls.Add(numDailyCapital);

            panelTop.Controls.Add(new Label { Text = "每次交易(張):", Location = new Point(460, 20), AutoSize = true });
            numTradeLots = new NumericUpDown
            {
                Location = new Point(550, 18),
                Width = 80,
                Maximum = 100m,
                Minimum = 1m,
                Value = 1m,
                Increment = 1m
            };
            panelTop.Controls.Add(numTradeLots);

            btnStart = new Button
            {
                Text = "啟動 AI 當沖模擬",
                Location = new Point(20, 65),
                Width = 160,
                Height = 40,
                BackColor = Color.LightGreen
            };
            btnStart.Click += BtnStart_Click;
            panelTop.Controls.Add(btnStart);

            btnStop = new Button
            {
                Text = "停止模擬",
                Location = new Point(190, 65),
                Width = 100,
                Height = 40,
                BackColor = Color.LightPink,
                Enabled = false
            };
            btnStop.Click += BtnStop_Click;
            panelTop.Controls.Add(btnStop);

            lblStatus = new Label
            {
                Text = "狀態: 尚未啟動",
                Location = new Point(310, 72),
                AutoSize = true,
                Font = new Font("微軟正黑體", 10, FontStyle.Bold)
            };
            panelTop.Controls.Add(lblStatus);

            lblPosition = new Label
            {
                Text = "等待啟動...",
                Location = new Point(20, 110),
                AutoSize = true,
                Font = new Font("Consolas", 9, FontStyle.Bold),
                ForeColor = Color.DarkBlue
            };
            panelTop.Controls.Add(lblPosition);

            this.Controls.Add(panelTop);

            rtbLog = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 10),
                BackColor = Color.Black,
                ForeColor = Color.LightGray
            };
            this.Controls.Add(rtbLog);
        }

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            _engine.Symbol = txtSymbol.Text.Trim();
            _engine.DailyCapitalLimit = numDailyCapital.Value;
            _engine.TradeLots = (int)numTradeLots.Value;

            btnStart.Enabled = false; btnStop.Enabled = true;
            txtSymbol.Enabled = false; numDailyCapital.Enabled = false; numTradeLots.Enabled = false;
            rtbLog.Clear();

            LogMessage("系統", "════════════════════════════════════════");
            LogMessage("系統", "台股當沖 AI 模擬引擎（修正版）啟動");
            LogMessage("系統", $"標的: {_engine.Symbol} | 本金: {_engine.DailyCapitalLimit:C0} | 每次: {_engine.TradeLots} 張");
            LogMessage("系統", "手續費 0.1425% + 當沖證交稅 0.15% | 漲跌停 ±10%");
            LogMessage("系統", "════════════════════════════════════════");

            await _engine.StartSimulationAsync();
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            _engine.StopSimulation();
            btnStart.Enabled = true; btnStop.Enabled = false;
            txtSymbol.Enabled = true; numDailyCapital.Enabled = true; numTradeLots.Enabled = true;
            LogMessage("系統", "AI 當沖模擬已停止。");

            if (lblStatus.InvokeRequired) lblStatus.Invoke(new Action(() => lblStatus.Text = "狀態: 已停止"));
            else lblStatus.Text = "狀態: 已停止";
        }

        private void Engine_OnLogMessage(string type, string message) => LogMessage(type, message);

        private void Engine_OnStatusUpdate(decimal availCap, decimal lockCap, int availLots, int lockLots, decimal price, bool isMarketOpen, decimal totalFee)
        {
            if (InvokeRequired) { Invoke(new Action(() => Engine_OnStatusUpdate(availCap, lockCap, availLots, lockLots, price, isMarketOpen, totalFee))); return; }

            string marketStr = isMarketOpen ? "🟢 交易中" : "🔴 休市中";
            lblStatus.Text = $"狀態: {marketStr}";
            lblPosition.Text =
                $"最新價: {price:F2} | " +
                $"可用庫存: {availLots} 張 (圈存 {lockLots} 張) | " +
                $"可用資金: {availCap:C0} (圈存 {lockCap:C0}) | " +
                $"累積手續費+稅: {totalFee:C0}";
        }

        private void LogMessage(string type, string message)
        {
            if (InvokeRequired) { Invoke(new Action(() => LogMessage(type, message))); return; }

            string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            rtbLog.SelectionStart = rtbLog.TextLength;
            rtbLog.SelectionLength = 0;
            rtbLog.SelectionColor = type switch
            {
                "AI決策" => Color.Cyan,
                "掛單" => Color.Yellow,
                "成交" => Color.Lime,
                "強制平倉" => Color.OrangeRed,
                "費用" => Color.Plum,
                "警告" => Color.Red,
                "行情" => Color.White,
                _ => Color.LightGray
            };
            rtbLog.AppendText($"[{timestamp}] [{type}] {message}\n");
            rtbLog.ScrollToCaret();
        }
    }

    // ════════════════════════════════════════════════════════════════
    //  AIAutoTradeEngine  ─  核心引擎（完整修正版）
    // ════════════════════════════════════════════════════════════════
    public class AIAutoTradeEngine
    {
        // ── 交易參數 ──────────────────────────────────────────────
        public string Symbol { get; set; } = "2330";
        public decimal DailyCapitalLimit { get; set; } = 250_000m;
        public int TradeLots { get; set; } = 1;   // 預設 1 張（台股最小單位）

        // ── 手續費常數 ─────────────────────────────────────────────
        private const decimal CommissionRate = 0.001425m; // 0.1425%
        private const decimal CommissionMinimum = 20m;       // 最低 20 元
        private const decimal DayTradeTaxRate = 0.0015m;   // 當沖證交稅 0.15%（一般 0.3%）

        // ── 漲跌停 ──────────────────────────────────────────────
        private const decimal DailyLimitPct = 0.10m;           // ±10%

        // ── 狀態欄位（全部改為「張數」，不再用股數）──────────────
        private bool _isRunning = false;
        private decimal _availableCapital;
        private decimal _lockedCapital;
        private int _availableLots;   // ✅ 修正7：直接存張數
        private int _lockedLots;      // ✅ 修正7：直接存張數

        // 當日買進張數（用於 13:25 強制平倉）
        private int _dayTradeBoughtLots = 0;

        private decimal _currentPrice;
        private decimal _openPrice;       // 當日開盤價（計算漲跌停用）
        private decimal _totalFee = 0m;   // 累積手續費+稅

        // ✅ 修正1：所有狀態變更都在此 lock 內進行
        private readonly object _stateLock = new object();

        // ✅ 修正2：SemaphoreSlim(1) 確保同時只有一筆訂單執行
        private readonly SemaphoreSlim _orderSemaphore = new SemaphoreSlim(1, 1);

        public event Action<string, string> OnLogMessage;
        public event Action<decimal, decimal, int, int, decimal, bool, decimal> OnStatusUpdate;

        private readonly Random _rnd = new Random();

        // ── 技術指標計算所需的滾動價格歷史 ────────────────────────────
        // 用於計算快速 EMA(9) 和慢速 EMA(21) 及 RSI(14)，取代純亂數信號
        private readonly Queue<decimal> _priceHistory = new Queue<decimal>();
        private const int FastEmaPeriod = 9;
        private const int SlowEmaPeriod = 21;
        private const int RsiPeriod = 14;
        private const int MinHistoryForSignal = SlowEmaPeriod + 1;  // 至少需要22根

        // ════════════════════════════════════════════════════════════
        //  啟動 / 停止
        // ════════════════════════════════════════════════════════════
        public async Task StartSimulationAsync()
        {
            _isRunning = true;
            _availableCapital = DailyCapitalLimit;
            _lockedCapital = 0m;
            _availableLots = 0;
            _lockedLots = 0;
            _dayTradeBoughtLots = 0;
            _totalFee = 0m;
            _currentPrice = 61.03m;
            _openPrice = _currentPrice;  // 開盤價 = 初始價

            TriggerStatusUpdate();

            await Task.Run(async () =>
            {
                while (_isRunning)
                {
                    SimulateMarketTick();

                    // ✅ 修正5：13:25 強制平倉檢查
                    if (IsForceCloseTime())
                    {
                        await ForceCloseAllPositionsAsync();
                        _isRunning = false;
                        OnLogMessage?.Invoke("系統", "13:25 強制平倉完成，當日交易結束。");
                        break;
                    }

                    // ✅ 修正4：只在交易時段決策
                    if (IsMarketOpen())
                    {
                        // ✅ 修正2：await 等待完成，不 fire-and-forget
                        await PerformAIDecisionAsync();
                    }

                    await Task.Delay(2000);
                }
            });
        }

        public void StopSimulation() => _isRunning = false;

        // ════════════════════════════════════════════════════════════
        //  修正3：股價模擬加入漲跌停 + 最低價保護
        // ════════════════════════════════════════════════════════════
        private void SimulateMarketTick()
        {
            // 每 tick ±0.3%（更接近真實 5 秒 tick 振幅）
            double changePct = (_rnd.NextDouble() * 0.6 - 0.3) / 100.0;
            decimal newPrice = _currentPrice * (1m + (decimal)changePct);

            // ✅ 修正3a：漲跌停限制 ±10%
            decimal upperLimit = Math.Round(_openPrice * (1m + DailyLimitPct), 2);
            decimal lowerLimit = Math.Round(_openPrice * (1m - DailyLimitPct), 2);
            newPrice = Math.Max(lowerLimit, Math.Min(upperLimit, newPrice));

            // ✅ 修正3b：最低價 0.01 元
            newPrice = Math.Max(0.01m, Math.Round(newPrice, 2));
            _currentPrice = newPrice;

            TriggerStatusUpdate();
            OnLogMessage?.Invoke("行情", $"{Symbol} 成交價: {_currentPrice:F2}  漲停: {upperLimit:F2}  跌停: {lowerLimit:F2}");
        }

        // ════════════════════════════════════════════════════════════
        //  修正1 + 修正2：AI 決策（技術指標信號取代純亂數）
        //  信號邏輯：快速 EMA(9) vs 慢速 EMA(21) 交叉 + RSI(14) 過濾
        //    買進：快線 > 慢線（上升趨勢）且 RSI < 65（未超買）
        //    賣出：快線 < 慢線（下降趨勢）且 RSI > 35（未超賣）
        // ════════════════════════════════════════════════════════════
        private async Task PerformAIDecisionAsync()
        {
            await Task.Delay(_rnd.Next(200, 800));

            // 更新滾動價格歷史（最多保留 SlowEmaPeriod+5 根）
            _priceHistory.Enqueue(_currentPrice);
            while (_priceHistory.Count > SlowEmaPeriod + 5)
                _priceHistory.Dequeue();

            // 資料不足時跳過（等待暖機期結束）
            if (_priceHistory.Count < MinHistoryForSignal) return;

            decimal[] prices = _priceHistory.ToArray();
            double fastEma = CalcEma(prices, FastEmaPeriod);
            double slowEma = CalcEma(prices, SlowEmaPeriod);
            double rsi = CalcRsi(prices, RsiPeriod);
            double priceVsOpen = _openPrice > 0
                ? (double)((_currentPrice - _openPrice) / _openPrice)
                : 0;

            // 信號判斷
            bool buySignal = fastEma > slowEma          // 上升趨勢
                             && rsi < 65                // 未超買
                             && priceVsOpen > -0.05;    // 未大幅跌離開盤 (-5%)
            bool sellSignal = fastEma < slowEma         // 下降趨勢
                              && rsi > 35               // 未超賣
                              && priceVsOpen < 0.05;    // 未大幅漲離開盤 (+5%)

            OnLogMessage?.Invoke("AI決策",
                $"EMA快({fastEma:F2}) EMA慢({slowEma:F2}) RSI({rsi:F1}) → " +
                (buySignal ? "📈買進" : sellSignal ? "📉賣出" : "⏸持觀望"));

            if (buySignal)
            {
                int lotsToTry = TradeLots;
                decimal costPerLot = _currentPrice * 1000m;

                // ✅ 修正1：在 lock 內同時計算 + 圈存，避免 TOCTOU
                int actualBuyLots = 0;
                lock (_stateLock)
                {
                    int maxAffordableLots = (int)(_availableCapital / costPerLot);
                    actualBuyLots = Math.Min(maxAffordableLots, lotsToTry);

                    if (actualBuyLots > 0)
                    {
                        decimal totalCost = actualBuyLots * costPerLot;
                        _availableCapital -= totalCost;
                        _lockedCapital += totalCost;
                    }
                }

                if (actualBuyLots > 0)
                {
                    OnLogMessage?.Invoke("AI決策",
                        $"[買進] 股價 {_currentPrice:F2}（一張 {costPerLot:C0}），買 {actualBuyLots} 張");
                    await ExecuteBuyAsync(_currentPrice, actualBuyLots);
                }
            }
            else if (sellSignal)
            {
                int actualSellLots = 0;

                // ✅ 修正1：在 lock 內判斷庫存並圈存
                lock (_stateLock)
                {
                    actualSellLots = Math.Min(_availableLots, TradeLots);
                    if (actualSellLots > 0)
                    {
                        _availableLots -= actualSellLots;
                        _lockedLots += actualSellLots;
                    }
                }

                if (actualSellLots > 0)
                {
                    OnLogMessage?.Invoke("AI決策",
                        $"[賣出] 股價 {_currentPrice:F2}，賣 {actualSellLots} 張");
                    await ExecuteSellAsync(_currentPrice, actualSellLots);
                }
            }
        }

        // ── EMA 計算（指數移動平均，種子取最早 period 根的簡單平均）──
        private static double CalcEma(decimal[] prices, int period)
        {
            if (prices.Length < period) return (double)prices[prices.Length - 1];
            // 種子：最早 period 根的簡單平均
            double ema = 0;
            int startIdx = prices.Length - period;
            for (int i = startIdx - Math.Min(period - 1, startIdx); i < startIdx; i++)
                ema += (double)prices[i];
            ema = prices.Length >= period * 2
                ? ema / (period - 1 > 0 ? period - 1 : 1)
                : (double)prices[startIdx];
            // 累積 EMA
            double k = 2.0 / (period + 1);
            for (int i = startIdx; i < prices.Length; i++)
                ema = (double)prices[i] * k + ema * (1 - k);
            return ema;
        }

        // ── RSI 計算（Wilder 平滑，period 根）─────────────────────────
        private static double CalcRsi(decimal[] prices, int period)
        {
            if (prices.Length < period + 1) return 50; // 資料不足回傳中性值
            int start = prices.Length - period - 1;
            double gains = 0, losses = 0;
            for (int i = start + 1; i <= start + period; i++)
            {
                double chg = (double)(prices[i] - prices[i - 1]);
                if (chg > 0) gains += chg; else losses -= chg;
            }
            double avgGain = gains / period;
            double avgLoss = losses / period;
            if (avgLoss == 0) return 100;
            return 100 - 100.0 / (1 + avgGain / avgLoss);
        }

        // ════════════════════════════════════════════════════════════
        //  買進流程（已在呼叫前完成圈存，此處只做撮合+確認）
        // ════════════════════════════════════════════════════════════
        private async Task ExecuteBuyAsync(decimal orderPrice, int lots)
        {
            // ✅ 修正2：確保同時只有一筆訂單
            await _orderSemaphore.WaitAsync();
            try
            {
                string orderId = "B" + DateTime.Now.ToString("HHmmssfff");
                decimal costPerLot = orderPrice * 1000m;
                decimal totalOrderAmount = lots * costPerLot;

                OnLogMessage?.Invoke("掛單",
                    $"[{orderId}] 買進委託 {lots} 張 | 掛單價 {orderPrice:F2} | 金額 {totalOrderAmount:C0}");

                await Task.Delay(_rnd.Next(300, 1000)); // 模擬送至交易所
                OnLogMessage?.Invoke("掛單", $"[{orderId}] 已進入撮合簿");

                await Task.Delay(_rnd.Next(500, 2000)); // 模擬等待成交

                // ✅ 修正8：加入滑價（買進可能貴 0~0.05%）
                decimal slippage = (decimal)(_rnd.NextDouble() * 0.0005);
                decimal fillPrice = Math.Round(orderPrice * (1m + slippage), 2);
                decimal fillAmount = lots * fillPrice * 1000m;

                // ✅ 修正6：計算手續費
                decimal commission = Math.Max(CommissionMinimum, fillAmount * CommissionRate);

                lock (_stateLock)
                {
                    // 解除圈存（原以 orderPrice 圈）並用真實成交金額調整
                    _lockedCapital -= totalOrderAmount;
                    // 成交金額差額退回（若滑價讓成交金額 > 圈存金額，再從可用扣）
                    decimal diff = fillAmount - totalOrderAmount;
                    _availableCapital -= diff;
                    _availableCapital -= commission; // 扣手續費
                    _totalFee += commission;

                    // 入庫（張數）+ 記錄當沖買進
                    _availableLots += lots;
                    _dayTradeBoughtLots += lots;
                }

                OnLogMessage?.Invoke("成交",
                    $"[{orderId}] 買進成交 {lots} 張 | 成交價 {fillPrice:F2} | 金額 {fillAmount:C0}");
                OnLogMessage?.Invoke("費用",
                    $"[{orderId}] 手續費 {commission:C0}（{CommissionRate:P4}）");

                TriggerStatusUpdate();
            }
            finally
            {
                _orderSemaphore.Release();
            }
        }

        // ════════════════════════════════════════════════════════════
        //  賣出流程（已在呼叫前完成庫存圈存）
        // ════════════════════════════════════════════════════════════
        private async Task ExecuteSellAsync(decimal orderPrice, int lots)
        {
            await _orderSemaphore.WaitAsync();
            try
            {
                string orderId = "S" + DateTime.Now.ToString("HHmmssfff");
                decimal orderAmount = lots * orderPrice * 1000m;

                OnLogMessage?.Invoke("掛單",
                    $"[{orderId}] 賣出委託 {lots} 張 | 掛單價 {orderPrice:F2} | 金額 {orderAmount:C0}");

                await Task.Delay(_rnd.Next(300, 1000));
                OnLogMessage?.Invoke("掛單", $"[{orderId}] 已進入撮合簿");

                await Task.Delay(_rnd.Next(500, 2000));

                // ✅ 修正8：賣出滑價（可能賣便宜 0~0.05%）
                decimal slippage = (decimal)(_rnd.NextDouble() * 0.0005);
                decimal fillPrice = Math.Round(orderPrice * (1m - slippage), 2);
                decimal fillAmount = lots * fillPrice * 1000m;

                // ✅ 修正6：手續費 + 當沖證交稅（折半）
                decimal commission = Math.Max(CommissionMinimum, fillAmount * CommissionRate);
                decimal tax = fillAmount * DayTradeTaxRate;
                decimal totalCost = commission + tax;

                lock (_stateLock)
                {
                    _lockedLots -= lots;
                    _availableCapital += fillAmount - totalCost;
                    _totalFee += totalCost;
                }

                OnLogMessage?.Invoke("成交",
                    $"[{orderId}] 賣出成交 {lots} 張 | 成交價 {fillPrice:F2} | 入帳 {fillAmount - totalCost:C0}");
                OnLogMessage?.Invoke("費用",
                    $"[{orderId}] 手續費 {commission:C0} + 當沖稅 {tax:C0} = {totalCost:C0}");

                TriggerStatusUpdate();
            }
            finally
            {
                _orderSemaphore.Release();
            }
        }

        // ════════════════════════════════════════════════════════════
        //  修正5：13:25 強制平倉（當沖核心規則）
        // ════════════════════════════════════════════════════════════
        private async Task ForceCloseAllPositionsAsync()
        {
            int lotsToClose;
            lock (_stateLock)
            {
                lotsToClose = _availableLots;
                if (lotsToClose > 0)
                {
                    _availableLots -= lotsToClose;
                    _lockedLots += lotsToClose;
                }
            }

            if (lotsToClose <= 0) return;

            OnLogMessage?.Invoke("強制平倉",
                $"⚠️ 13:25 強制平倉！賣出庫存 {lotsToClose} 張（當沖規則）");

            await ExecuteSellAsync(_currentPrice, lotsToClose);
        }

        // ════════════════════════════════════════════════════════════
        //  修正4：交易時段判斷
        // ════════════════════════════════════════════════════════════
        private static bool IsMarketOpen()
        {
            var now = DateTime.Now;
            // 週一 ~ 週五（排除假日需額外處理）
            if (now.DayOfWeek == DayOfWeek.Saturday || now.DayOfWeek == DayOfWeek.Sunday)
                return false;

            var open = now.Date.AddHours(9).AddMinutes(0);
            var close = now.Date.AddHours(13).AddMinutes(30);
            return now >= open && now <= close;
        }

        private static bool IsForceCloseTime()
        {
            var now = DateTime.Now;
            if (now.DayOfWeek == DayOfWeek.Saturday || now.DayOfWeek == DayOfWeek.Sunday)
                return false;
            var forceClose = now.Date.AddHours(13).AddMinutes(25);
            var marketEnd = now.Date.AddHours(13).AddMinutes(30);
            return now >= forceClose && now <= marketEnd;
        }

        private void TriggerStatusUpdate()
        {
            int availLots, lockLots;
            decimal availCap, lockCap;
            lock (_stateLock)
            {
                availLots = _availableLots;
                lockLots = _lockedLots;
                availCap = _availableCapital;
                lockCap = _lockedCapital;
            }
            OnStatusUpdate?.Invoke(availCap, lockCap, availLots, lockLots, _currentPrice, IsMarketOpen(), _totalFee);
        }
    }
}