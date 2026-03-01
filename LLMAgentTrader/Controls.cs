using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace LLMAgentTrader
{
    public class GlassPanel : Panel
    {
        public int Radius { get; set; } = 15;
        public Color BaseColor { get; set; } = Color.FromArgb(30, 32, 40);
        public GlassPanel() { DoubleBuffered = true; BackColor = Color.Transparent; }
        protected override void OnPaint(PaintEventArgs e)
        {
            if (Width <= 0 || Height <= 0) return;
            base.OnPaint(e); e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = QuantChartPanel.GetRoundedRect(new Rectangle(0, 0, Width - 1, Height - 1), Radius))
            {
                using (var brush = new SolidBrush(BaseColor)) e.Graphics.FillPath(brush, path);
                using (var pen = new Pen(Color.FromArgb(40, 255, 255, 255), 1)) e.Graphics.DrawPath(pen, path);
            }
        }
    }

    public class GlassButton : Button
    {
        public Color BaseColor { get; set; } = Color.FromArgb(0, 120, 212);
        public GlassButton() { FlatStyle = FlatStyle.Flat; ForeColor = Color.White; BackColor = Color.Transparent; FlatAppearance.BorderSize = 0; }
        protected override void OnPaint(PaintEventArgs e)
        {
            if (Width <= 0 || Height <= 0) return;
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Color drawColor = ClientRectangle.Contains(PointToClient(MousePosition)) ? ControlPaint.Light(BaseColor) : BaseColor;
            using (var path = QuantChartPanel.GetRoundedRect(new Rectangle(0, 0, Width - 1, Height - 1), 10))
            {
                using (var brush = new SolidBrush(drawColor)) e.Graphics.FillPath(brush, path);
                TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, ForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
        }
    }

    public class OrderBookPanel : Panel
    {
        private TwseTick _tick;
        private string _fallbackMsg = "等待數據...";
        public OrderBookPanel() { DoubleBuffered = true; BackColor = Color.Transparent; }
        public void UpdateData(TwseTick tick) { _tick = tick; Invalidate(); }
        public void ClearData(string msg) { _tick = null; _fallbackMsg = msg; Invalidate(); }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
            float w = Width;
            g.DrawString("🇹🇼 TWSE 五檔", new Font("Segoe UI", 12, FontStyle.Bold), Brushes.LightGray, 10, 15);
            if (_tick == null) { g.DrawString(_fallbackMsg, new Font("Segoe UI", 10), Brushes.Gray, 10, 45); return; }
            g.DrawString($"同步: {_tick.Time}", new Font("Consolas", 9), Brushes.Gray, 10, 40);
            float yBase = 65, rowH = 28;
            for (int i = 4; i >= 0; i--)
            {
                double p = i < _tick.AskPrices.Count ? _tick.AskPrices[i] : 0;
                int v = i < _tick.AskVolumes.Count ? _tick.AskVolumes[i] : 0;
                Color c = p > _tick.YesterdayClose ? Color.SpringGreen : (p < _tick.YesterdayClose ? Color.LightCoral : Color.White);
                if (p == 0) c = Color.Gray;
                g.DrawString($"賣 {i + 1}", new Font("Segoe UI", 9.5F), Brushes.Gray, 10, yBase);
                g.DrawString(p > 0 ? $"{p:F2}" : "-", new Font("Consolas", 12, FontStyle.Bold), new SolidBrush(c), 60, yBase);
                g.DrawString(v > 0 ? $"{v}" : "-", new Font("Consolas", 11), Brushes.White, w - 10, yBase, new StringFormat { Alignment = StringAlignment.Far });
                yBase += rowH;
            }
            g.DrawLine(new Pen(Color.FromArgb(50, 255, 255, 255)), 10, yBase, w - 10, yBase); yBase += 8;
            for (int i = 0; i < 5; i++)
            {
                double p = i < _tick.BidPrices.Count ? _tick.BidPrices[i] : 0;
                int v = i < _tick.BidVolumes.Count ? _tick.BidVolumes[i] : 0;
                Color c = p > _tick.YesterdayClose ? Color.SpringGreen : (p < _tick.YesterdayClose ? Color.LightCoral : Color.White);
                if (p == 0) c = Color.Gray;
                g.DrawString($"買 {i + 1}", new Font("Segoe UI", 9.5F), Brushes.Gray, 10, yBase);
                g.DrawString(p > 0 ? $"{p:F2}" : "-", new Font("Consolas", 12, FontStyle.Bold), new SolidBrush(c), 60, yBase);
                g.DrawString(v > 0 ? $"{v}" : "-", new Font("Consolas", 11), Brushes.White, w - 10, yBase, new StringFormat { Alignment = StringAlignment.Far });
                yBase += rowH;
            }
            yBase += 10;
            g.DrawString($"總量: {_tick.Volume:N0}", new Font("Segoe UI", 11, FontStyle.Bold), Brushes.LightGoldenrodYellow, 10, yBase);
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  互動 K 線圖表引擎 (含 Fibonacci 回調線)
    // ────────────────────────────────────────────────────────────────────────────
    public class QuantChartPanel : Panel
    {
        private List<MarketData> _data = new List<MarketData>();
        private string _ticker = "";
        private string _companyName = "";
        private double _pe = 0;
        private double _yield = 0;

        // Fibonacci 開關與資料
        public bool ShowFibonacci { get; set; } = true;
        private List<FibLevel> _fibLevels = new List<FibLevel>();

        public bool IsIntraday { get; set; } = false;
        public Action<int> OnDataPointClicked;
        private int visStart = 0, visEnd = 0, lastMouseX = 0, hoveredIdx = -1, selectedIdx = -1;
        private bool isDragging = false;

        // Fibonacci 顏色對應 (0%, 23.6%, 38.2%, 50%, 61.8%, 78.6%, 100%)
        private static readonly Color[] FibColors =
        {
            Color.FromArgb(150, 200, 200, 0),   // 0%    黃
            Color.FromArgb(150, 0, 200, 200),   // 23.6% 青
            Color.FromArgb(150, 100, 180, 255), // 38.2% 藍
            Color.FromArgb(150, 255, 165, 0),   // 50%   橙
            Color.FromArgb(150, 100, 180, 255), // 61.8% 藍
            Color.FromArgb(150, 0, 200, 200),   // 78.6% 青
            Color.FromArgb(150, 200, 200, 0),   // 100%  黃
        };

        public QuantChartPanel() { DoubleBuffered = true; Cursor = Cursors.Cross; BackColor = Color.Transparent; }

        public void UpdateData(List<MarketData> d, string t, string companyName = "", double pe = 0, double yld = 0)
        {
            _data = d ?? new List<MarketData>();
            _ticker = t; _companyName = companyName; _pe = pe; _yield = yld;
            if (_data.Count > 0) { int def = 80; visStart = Math.Max(0, _data.Count - def); visEnd = _data.Count - 1; }
            else { visStart = 0; visEnd = 0; }
            selectedIdx = -1; hoveredIdx = -1;

            // 自動計算 Fibonacci
            if (ShowFibonacci && _data.Count >= 10)
                _fibLevels = FibonacciEngine.Calculate(_data, Math.Min(120, _data.Count));
            else
                _fibLevels.Clear();

            Invalidate();
        }

        public void SetSelectedIndex(int i)
        {
            selectedIdx = i;
            if (selectedIdx < visStart || selectedIdx > visEnd)
            {
                int range = visEnd - visStart;
                visStart = Math.Max(0, selectedIdx - range / 2);
                visEnd = Math.Min(_data.Count - 1, visStart + range);
            }
            Invalidate();
        }

        public static GraphicsPath GetRoundedRect(RectangleF rect, float radius)
        {
            var path = new GraphicsPath();
            if (rect.Width <= 0 || rect.Height <= 0) return path;
            float d = radius * 2;
            if (d <= 0) { path.AddRectangle(rect); return path; }
            d = Math.Min(d, Math.Min(rect.Width, rect.Height));
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            if (!Focused) Focus();
            if (_data.Count == 0) return;
            int zoom = Math.Max(1, (int)((visEnd - visStart) * 0.1));
            if (e.Delta > 0) { visStart += zoom; visEnd -= zoom; } else { visStart -= zoom; visEnd += zoom; }
            if (visEnd - visStart < 10) visEnd = visStart + 10;
            if (visStart < 0) visStart = 0;
            if (visEnd >= _data.Count) visEnd = _data.Count - 1;
            if (visStart >= visEnd) visStart = Math.Max(0, visEnd - 10);
            Invalidate();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e); Focus();
            if (e.Button == MouseButtons.Left && _data.Count > 0) { isDragging = true; lastMouseX = e.X; }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button == MouseButtons.Left)
            {
                if (isDragging && Math.Abs(e.X - lastMouseX) < 5 && hoveredIdx >= 0)
                    OnDataPointClicked?.Invoke(hoveredIdx);
                isDragging = false;
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_data.Count == 0) return;
            float step = (Width - 100f) / Math.Max(1, visEnd - visStart);
            if (isDragging)
            {
                int shift = -(int)((e.X - lastMouseX) / step);
                if (shift != 0)
                {
                    int range = visEnd - visStart;
                    visStart += shift; visEnd += shift;
                    if (visStart < 0) { visStart = 0; visEnd = range; }
                    if (visEnd >= _data.Count) { visEnd = _data.Count - 1; visStart = visEnd - range; }
                    lastMouseX = e.X; Invalidate();
                }
            }
            else
            {
                int newHover = visStart + (int)Math.Round((e.X - 50f) / step);
                newHover = Math.Max(visStart, Math.Min(visEnd, newHover));
                if (hoveredIdx != newHover) { hoveredIdx = newHover; Invalidate(); }
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (_data.Count < 2) return;
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var visData = _data.Skip(visStart).Take(visEnd - visStart + 1).ToList();
            if (visData.Count == 0) return;

            float px = 50, py = 60, w = Width - 100;
            float hTop = (Height - 110) * 0.75f, hBot = (Height - 110) * 0.2f;
            float pyBot = py + hTop + (Height - 110) * 0.05f;

            double maxP = visData.Max(x => Math.Max(x.High, Math.Max(x.BB_Upper > 0 ? x.BB_Upper : x.High, x.ResistanceLevel > 0 ? x.ResistanceLevel : x.High)));
            double minP = visData.Where(x => x.Low > 0).Min(x => Math.Min(x.Low, Math.Min(x.BB_Lower > 0 ? x.BB_Lower : x.Low, x.SupportLevel > 0 ? x.SupportLevel : x.Low)));
            if (Math.Abs(maxP - minP) < 0.0001) { maxP += 1; minP -= 1; }

            // Fibonacci 範圍也要納入
            if (_fibLevels.Count > 0)
            {
                double fibMax = _fibLevels.Max(f => f.Price);
                double fibMin = _fibLevels.Min(f => f.Price);
                maxP = Math.Max(maxP, fibMax);
                minP = Math.Min(minP, fibMin);
            }

            double rangeP = maxP - minP; maxP += rangeP * 0.05; minP -= rangeP * 0.05;
            double maxV = visData.Max(x => x.Volume); if (maxV == 0) maxV = 1;
            float step = w / Math.Max(1, visData.Count - 1);

            // ── 支撐壓力線 ─────────────────────────────────────────────────────
            using (var resPen = new Pen(Color.FromArgb(100, 100, 150, 255), 2) { DashStyle = DashStyle.Dot })
            using (var supPen = new Pen(Color.FromArgb(100, 200, 100, 255), 2) { DashStyle = DashStyle.Dot })
            {
                double currRes = visData.Last().ResistanceLevel, currSup = visData.Last().SupportLevel;
                if (currRes > 0)
                {
                    float y = py + hTop - (float)((currRes - minP) / (maxP - minP) * hTop);
                    g.DrawLine(resPen, px, y, px + w, y);
                    g.DrawString($"壓力 {currRes:F2}", new Font("Consolas", 8), Brushes.LightSkyBlue, px + 5, y - 15);
                }
                if (currSup > 0)
                {
                    float y = py + hTop - (float)((currSup - minP) / (maxP - minP) * hTop);
                    g.DrawLine(supPen, px, y, px + w, y);
                    g.DrawString($"支撐 {currSup:F2}", new Font("Consolas", 8), Brushes.MediumOrchid, px + 5, y + 5);
                }
            }

            // ── Fibonacci 回調線 ──────────────────────────────────────────────
            if (ShowFibonacci && _fibLevels.Count > 0)
            {
                for (int fi = 0; fi < _fibLevels.Count; fi++)
                {
                    var fib = _fibLevels[fi];
                    if (fib.Price <= 0) continue;
                    float fy = py + hTop - (float)((fib.Price - minP) / (maxP - minP) * hTop);
                    if (fy < py || fy > py + hTop) continue;

                    Color fibColor = fi < FibColors.Length ? FibColors[fi] : Color.FromArgb(120, 200, 200, 200);
                    using (var fibPen = new Pen(fibColor, 1) { DashStyle = DashStyle.Dash })
                        g.DrawLine(fibPen, px, fy, px + w, fy);

                    // 標籤 (靠右對齊)
                    string label = $"Fib {fib.Ratio:P1}  {fib.Price:F2}";
                    g.DrawString(label, new Font("Consolas", 7.5F), new SolidBrush(fibColor), px + w - 100, fy - 13);
                }
            }

            // ── K 棒與技術線 ──────────────────────────────────────────────────
            var bbUp = new List<PointF>(); var bbDn = new List<PointF>();
            var ema50 = new List<PointF>(); var vwapLine = new List<PointF>();

            for (int i = 0; i < visData.Count; i++)
            {
                float x = px + i * step;
                float vH = (float)(visData[i].Volume / maxV * hBot);
                g.FillRectangle(new SolidBrush(Color.FromArgb(70, 100, 150, 255)), x - 2, pyBot + hBot - vH, 4, vH);

                float yO = py + hTop - (float)((visData[i].Open - minP) / (maxP - minP) * hTop);
                float yC = py + hTop - (float)((visData[i].Close - minP) / (maxP - minP) * hTop);
                float yH = py + hTop - (float)((visData[i].High - minP) / (maxP - minP) * hTop);
                float yL = py + hTop - (float)((visData[i].Low - minP) / (maxP - minP) * hTop);
                Color c = visData[i].Close >= visData[i].Open ? Color.FromArgb(255, 80, 80) : Color.FromArgb(80, 255, 120);
                using (var pen = new Pen(c, 1.5f)) { g.DrawLine(pen, x, yH, x, yL); g.FillRectangle(new SolidBrush(c), x - 3, Math.Min(yO, yC), 6, Math.Max(1, Math.Abs(yO - yC))); }

                if (visData[i].BB_Upper > 0) { bbUp.Add(new PointF(x, py + hTop - (float)((visData[i].BB_Upper - minP) / (maxP - minP) * hTop))); bbDn.Add(new PointF(x, py + hTop - (float)((visData[i].BB_Lower - minP) / (maxP - minP) * hTop))); }
                if (visData[i].EMA_50 > 0) ema50.Add(new PointF(x, py + hTop - (float)((visData[i].EMA_50 - minP) / (maxP - minP) * hTop)));
                if (visData[i].VWAP > 0) vwapLine.Add(new PointF(x, py + hTop - (float)((visData[i].VWAP - minP) / (maxP - minP) * hTop)));

                if (visData[i].AgentAction == "Buy") g.FillPolygon(Brushes.SpringGreen, new[] { new PointF(x, yL + 5), new PointF(x - 5, yL + 15), new PointF(x + 5, yL + 15) });
                if (visData[i].AgentAction == "Sell") g.FillPolygon(Brushes.LightCoral, new[] { new PointF(x, yH - 5), new PointF(x - 5, yH - 15), new PointF(x + 5, yH - 15) });
            }

            if (bbUp.Count > 1) { using (var p = new Pen(Color.FromArgb(60, 255, 255, 255)) { DashStyle = DashStyle.Dash }) { g.DrawLines(p, bbUp.ToArray()); g.DrawLines(p, bbDn.ToArray()); } }
            if (ema50.Count > 1) g.DrawLines(new Pen(Color.Orange, 1.5f), ema50.ToArray());
            if (vwapLine.Count > 1) g.DrawLines(new Pen(Color.Yellow, 2f), vwapLine.ToArray());

            // 最高最低標記
            var validBars = visData.Where(x => x.High > 0 && x.Low > 0).ToList();
            if (validBars.Count > 0)
            {
                double absH = validBars.Max(x => x.High), absL = validBars.Min(x => x.Low);
                int hiIdx = visData.FindIndex(x => x.High == absH), loIdx = visData.FindIndex(x => x.Low == absL);
                if (hiIdx >= 0) { float hX = px + hiIdx * step, hY = py + hTop - (float)((absH - minP) / (maxP - minP) * hTop); g.DrawString($"{absH:F2}", new Font("Consolas", 10, FontStyle.Bold), Brushes.LightCoral, hX - 15, hY - 20); }
                if (loIdx >= 0) { float lX = px + loIdx * step, lY = py + hTop - (float)((absL - minP) / (maxP - minP) * hTop); g.DrawString($"{absL:F2}", new Font("Consolas", 10, FontStyle.Bold), Brushes.SpringGreen, lX - 15, lY + 5); }
            }

            // 標題
            var titleFont = new Font("Microsoft JhengHei UI", 14, FontStyle.Bold);
            var fundFont = new Font("Consolas", 11, FontStyle.Bold);
            float titleY = 5;
            string titleStr = $"{_companyName} ({_ticker})";
            g.DrawString(titleStr, titleFont, Brushes.Gold, px, titleY);
            float titleW = g.MeasureString(titleStr, titleFont).Width;
            string peStr = _pe > 0 ? _pe.ToString("F2") : "-", yldStr = _yield > 0 ? _yield.ToString("P2") : "-";
            g.DrawString($"P/E: {peStr}  殖利率: {yldStr}", fundFont, Brushes.LightSkyBlue, px + titleW + 15, titleY + 3);

            // OHLCV 顯示
            var disp = (hoveredIdx >= visStart && hoveredIdx <= visEnd) ? _data[hoveredIdx] : visData.Last();
            var topFont = new Font("Segoe UI", 10.5F, FontStyle.Bold); float cX = px, ohlcY = 32;
            g.DrawString("開", topFont, Brushes.Gray, cX, ohlcY); cX += 20; g.DrawString($"{disp.Open:F2}", topFont, Brushes.White, cX, ohlcY); cX += 60;
            g.DrawString("高", topFont, Brushes.Gray, cX, ohlcY); cX += 20; g.DrawString($"{disp.High:F2}", topFont, Brushes.LightCoral, cX, ohlcY); cX += 60;
            g.DrawString("低", topFont, Brushes.Gray, cX, ohlcY); cX += 20; g.DrawString($"{disp.Low:F2}", topFont, Brushes.SpringGreen, cX, ohlcY); cX += 60;
            g.DrawString("收", topFont, Brushes.Gray, cX, ohlcY); cX += 20;
            Color curC = disp.Close >= disp.Open ? Color.LightCoral : Color.SpringGreen;
            g.DrawString($"{disp.Close:F2}", topFont, new SolidBrush(curC), cX, ohlcY); cX += 60;
            g.DrawString("量", topFont, Brushes.Gray, cX, ohlcY); cX += 25; g.DrawString($"{disp.Volume / 1000:N0}K", topFont, Brushes.White, cX, ohlcY); cX += 60;
            if (disp.VWAP > 0) { g.DrawString("VWAP", topFont, Brushes.Gold, cX, ohlcY); cX += 45; g.DrawString($"{disp.VWAP:F2}", topFont, Brushes.White, cX, ohlcY); }

            // 十字線
            if (hoveredIdx >= visStart && hoveredIdx <= visEnd)
            {
                float hx = px + (hoveredIdx - visStart) * step;
                float hy = py + hTop - (float)((_data[hoveredIdx].Close - minP) / (maxP - minP) * hTop);
                using (var crossPen = new Pen(Color.FromArgb(100, 255, 255, 255)))
                { g.DrawLine(crossPen, hx, py, hx, py + hTop + hBot); g.DrawLine(crossPen, px, hy, px + w, hy); }

                string dateStr = IsIntraday ? _data[hoveredIdx].Date.ToString("HH:mm:ss") : _data[hoveredIdx].Date.ToString("MM/dd");
                string tip = $"{dateStr}\nC: {_data[hoveredIdx].Close:F2}\nV: {_data[hoveredIdx].Volume / 1000:N0}K";
                g.FillRectangle(new SolidBrush(Color.FromArgb(230, 25, 28, 35)), hx + 8, hy - 45, 95, 50);
                g.DrawRectangle(new Pen(Color.FromArgb(50, 255, 255, 255)), hx + 8, hy - 45, 95, 50);
                g.DrawString(tip, new Font("Consolas", 8.5F), Brushes.White, hx + 12, hy - 40);
            }
            if (selectedIdx >= visStart && selectedIdx <= visEnd)
            {
                float sx = px + (selectedIdx - visStart) * step;
                g.DrawLine(new Pen(Color.DeepSkyBlue, 2), sx, py, sx, py + hTop);
            }
        }
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  浮動聊天視窗
    // ────────────────────────────────────────────────────────────────────────────
    public class FloatingChatForm : Form
    {
        private RichTextBox rtbChat;
        private TextBox txtInput;
        private GlassButton btnSend;
        private string apiKey;
        private string ticker;
        private Func<string> getCurrentContext;
        private List<object> chatMessages = new List<object>();

        public FloatingChatForm(string apiKey, string ticker, Func<string> getContextFunc)
        {
            this.apiKey = apiKey; this.ticker = ticker; this.getCurrentContext = getContextFunc;
            Text = "💬 N-of-1 智能交易助理";
            Size = new Size(500, 650);
            BackColor = Color.FromArgb(20, 22, 30);
            ForeColor = Color.White;
            TopMost = true;
            StartPosition = FormStartPosition.CenterScreen;

            var pnlMain = new GlassPanel { Dock = DockStyle.Fill, Radius = 0, BaseColor = Color.Transparent, Padding = new Padding(15) };
            rtbChat = new RichTextBox { Dock = DockStyle.Fill, BackColor = Color.FromArgb(15, 18, 25), ForeColor = Color.White, BorderStyle = BorderStyle.None, Font = new Font("Segoe UI", 11F), ReadOnly = true };
            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 60, Padding = new Padding(0, 15, 0, 0) };
            var pnlInput = new GlassPanel { Dock = DockStyle.Fill, Radius = 10, BaseColor = Color.FromArgb(35, 38, 45), Padding = new Padding(10, 5, 10, 5) };
            txtInput = new TextBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Color.FromArgb(35, 38, 45), ForeColor = Color.White, Font = new Font("Segoe UI", 12F) };
            txtInput.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) { ev.SuppressKeyPress = true; btnSend.PerformClick(); } };
            pnlInput.Controls.Add(txtInput);
            btnSend = new GlassButton { Text = "發送", Dock = DockStyle.Right, Width = 90, BaseColor = Color.FromArgb(0, 120, 212), Margin = new Padding(10, 0, 0, 0), Font = new Font("Segoe UI", 10F, FontStyle.Bold) };
            btnSend.Click += BtnSend_Click;
            pnlBottom.Controls.Add(pnlInput);
            pnlBottom.Controls.Add(new Panel { Dock = DockStyle.Right, Width = 10 });
            pnlBottom.Controls.Add(btnSend);
            pnlMain.Controls.Add(rtbChat); pnlMain.Controls.Add(pnlBottom);
            Controls.Add(pnlMain);
            AppendChat("系統", $"歡迎使用對話助理！點選主視窗數據可帶入上下文。({ticker})", Color.Gray);
        }

        private async void BtnSend_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtInput.Text)) return;
            string userMsg = txtInput.Text; txtInput.Clear();
            AppendChat("你", userMsg, Color.LightSkyBlue);
            ExcelChatLogger.LogMessage(ticker, "User", userMsg);
            btnSend.Enabled = false;
            try
            {
                if (chatMessages.Count == 0) chatMessages.Add(new { role = "system", content = "你是一位量化交易顧問。🚨請務必使用「繁體中文」回答。" });
                string ctx = getCurrentContext();
                chatMessages.Add(new { role = "user", content = $"[{ctx}]\n使用者問: {userMsg}" });
                AppendChat("AI", "", Color.SpringGreen);
                var sb = new StringBuilder();
                await LLMService.StreamChat(apiKey, chatMessages, (chunk) => Invoke(new Action(() => { rtbChat.AppendText(chunk); rtbChat.SelectionStart = rtbChat.TextLength; rtbChat.ScrollToCaret(); sb.Append(chunk); })));
                rtbChat.AppendText("\r\n\r\n");
                ExcelChatLogger.LogMessage(ticker, "AI", sb.ToString());
            }
            catch (Exception ex) { AppendChat("系統", $"錯誤: {ex.Message}", Color.LightCoral); }
            finally { btnSend.Enabled = true; txtInput.Focus(); }
        }

        private void AppendChat(string name, string msg, Color color)
        {
            rtbChat.SelectionStart = rtbChat.TextLength;
            rtbChat.SelectionFont = new Font(rtbChat.Font, FontStyle.Bold); rtbChat.SelectionColor = color;
            rtbChat.AppendText($"[{name}] ");
            rtbChat.SelectionFont = new Font(rtbChat.Font, FontStyle.Regular); rtbChat.SelectionColor = Color.White;
            rtbChat.AppendText($"{msg}\r\n"); rtbChat.ScrollToCaret();
        }
    }

    // ── 績效曲線圖控制項 ─────────────────────────────────────────────────────────
    public class EquityCurvePanel : Panel
    {
        private List<(DateTime date, double cumPnl)> _points = new();
        public EquityCurvePanel() { DoubleBuffered = true; }

        public void UpdateData(List<TradeJournalEntry> entries)
        {
            _points.Clear();
            double cumPnl = 0;
            foreach (var e in entries.OrderBy(e => e.TradeDate))
            {
                cumPnl += e.PnL;
                _points.Add((e.TradeDate, cumPnl));
            }
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            int pad = 30;
            var rect = new Rectangle(pad, 4, Width - pad - 8, Height - 14);
            if (rect.Width <= 0 || rect.Height <= 0) return;

            // 背景格線
            using (var gpen = new Pen(Color.FromArgb(30, 255, 255, 255), 1))
            {
                for (int i = 0; i <= 4; i++)
                {
                    int y = rect.Top + rect.Height * i / 4;
                    g.DrawLine(gpen, rect.Left, y, rect.Right, y);
                }
            }

            if (_points.Count < 2)
            {
                TextRenderer.DrawText(g, "尚無已平倉交易資料", new Font("Segoe UI", 9F),
                    rect, Color.FromArgb(80, 90, 110), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                return;
            }

            double minPnl = _points.Min(p => p.cumPnl);
            double maxPnl = _points.Max(p => p.cumPnl);
            if (Math.Abs(maxPnl - minPnl) < 0.01) maxPnl = minPnl + 1;

            float ScaleX(int idx) => rect.Left + (float)idx / (_points.Count - 1) * rect.Width;
            float ScaleY(double pnl) => rect.Bottom - (float)((pnl - minPnl) / (maxPnl - minPnl)) * rect.Height;

            // 填充漸層
            var pts = new PointF[_points.Count + 2];
            for (int i = 0; i < _points.Count; i++)
                pts[i] = new PointF(ScaleX(i), ScaleY(_points[i].cumPnl));
            pts[_points.Count] = new PointF(ScaleX(_points.Count - 1), rect.Bottom);
            pts[_points.Count + 1] = new PointF(ScaleX(0), rect.Bottom);

            using (var fillBrush = new LinearGradientBrush(
                new PointF(0, rect.Top), new PointF(0, rect.Bottom),
                Color.FromArgb(60, 0, 200, 100), Color.Transparent))
                g.FillPolygon(fillBrush, pts);

            // 曲線
            var linePts = pts.Take(_points.Count).ToArray();
            using (var pen = new Pen(Color.FromArgb(0, 220, 120), 2.5F))
                g.DrawLines(pen, linePts);

            // 零線
            if (minPnl < 0 && maxPnl > 0)
            {
                float yZero = ScaleY(0);
                using (var zeroPen = new Pen(Color.FromArgb(80, 255, 100, 60), 1) { DashStyle = DashStyle.Dash })
                    g.DrawLine(zeroPen, rect.Left, yZero, rect.Right, yZero);
            }

            // 最終損益標籤
            double last = _points.Last().cumPnl;
            Color lblColor = last >= 0 ? Color.SpringGreen : Color.LightCoral;
            string lblText = $"{(last >= 0 ? "+" : "")}{last:N0}";
            TextRenderer.DrawText(g, lblText, new Font("Consolas", 10F, FontStyle.Bold),
                new Point((int)ScaleX(_points.Count - 1) - 50, rect.Top + 2), lblColor);

            // Y 軸標籤
            TextRenderer.DrawText(g, $"{maxPnl:N0}", new Font("Segoe UI", 8F),
                new Point(0, rect.Top), Color.FromArgb(120, 130, 150), TextFormatFlags.Default);
            TextRenderer.DrawText(g, $"{minPnl:N0}", new Font("Segoe UI", 8F),
                new Point(0, rect.Bottom - 16), Color.FromArgb(120, 130, 150), TextFormatFlags.Default);
        }
    }
}