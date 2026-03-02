using System.Drawing;

namespace LLMAgentTrader
{
    // ════════════════════════════════════════════════════════════════════════
    //  Alpha-Twin Design System – 統一配色 / 字型 / 尺寸常數
    // ════════════════════════════════════════════════════════════════════════

    /// <summary>全域色彩 Token（所有 UI 皆從此取色，禁止散落的 FromArgb 硬編碼）</summary>
    public static class ThemeColors
    {
        // ── Backgrounds ──────────────────────────────────────────────────────
        public static readonly Color Background    = Color.FromArgb(10, 12, 18);    // 最深底層
        public static readonly Color Surface       = Color.FromArgb(18, 20, 28);    // Panel 背景
        public static readonly Color SurfaceAlt    = Color.FromArgb(22, 25, 36);    // 交替行 / 輔助面
        public static readonly Color Card          = Color.FromArgb(28, 32, 42);    // 卡片 / 標題列
        public static readonly Color NavBackground = Color.FromArgb(13, 15, 24);    // 左側導覽

        // ── Borders & Dividers ───────────────────────────────────────────────
        public static readonly Color Border        = Color.FromArgb(38, 44, 62);
        public static readonly Color Divider       = Color.FromArgb(30, 36, 52);

        // ── Inputs ───────────────────────────────────────────────────────────
        public static readonly Color Input         = Color.FromArgb(34, 38, 52);    // 輸入框正常
        public static readonly Color InputFocus    = Color.FromArgb(38, 56, 88);    // 輸入框 focus
        public static readonly Color InputBorder   = Color.FromArgb(55, 65, 90);    // 未 focus 框線顏色

        // ── Text ─────────────────────────────────────────────────────────────
        public static readonly Color TextPrimary   = Color.White;
        public static readonly Color TextSecondary = Color.FromArgb(160, 168, 188); // label 次要
        public static readonly Color TextMuted     = Color.FromArgb(95, 105, 128);  // hint / caption

        // ── Accent & Actions ──────────────────────────────────────────────────
        public static readonly Color Accent        = Color.FromArgb(0, 122, 218);   // 主要動作 (藍)
        public static readonly Color AccentHover   = Color.FromArgb(28, 152, 255);  // hover 亮藍
        public static readonly Color AccentAlt     = Color.FromArgb(0, 155, 100);   // 次要動作 (綠)
        public static readonly Color AccentAltHov  = Color.FromArgb(0, 190, 120);
        public static readonly Color Danger        = Color.FromArgb(220, 55, 72);   // 危險 / 刪除
        public static readonly Color DangerHover   = Color.FromArgb(255, 75, 90);
        public static readonly Color Gold          = Color.FromArgb(220, 170, 0);   // 提示 / Kelly

        // ── Status ───────────────────────────────────────────────────────────
        public static readonly Color StatusOk      = Color.FromArgb(0, 215, 110);   // 正面
        public static readonly Color StatusWarn    = Color.FromArgb(255, 160, 0);   // 警告
        public static readonly Color StatusErr     = Color.FromArgb(240, 65, 85);   // 錯誤
        public static readonly Color StatusInfo    = Color.FromArgb(100, 170, 255); // 資訊

        // ── Nav ──────────────────────────────────────────────────────────────
        public static readonly Color NavHover      = Color.FromArgb(28, 35, 55);
        public static readonly Color NavActive     = Color.FromArgb(22, 45, 80);
        public static readonly Color NavActiveLine = Color.FromArgb(0, 140, 255);   // 左側藍線

        // ── DGV Grid ─────────────────────────────────────────────────────────
        public static readonly Color GridHeader    = Color.FromArgb(22, 26, 38);
        public static readonly Color GridRow       = Color.FromArgb(18, 20, 28);
        public static readonly Color GridRowAlt    = Color.FromArgb(22, 25, 36);
        public static readonly Color GridSelect    = Color.FromArgb(35, 78, 155);
        public static readonly Color GridLine      = Color.FromArgb(32, 38, 55);
    }

    /// <summary>字型常數（所有 UI 從此取字型，禁止分散的 new Font() 硬編碼）</summary>
    public static class AppFonts
    {
        // ── 標題系列 ─────────────────────────────────────────────────────────
        public static readonly Font H1        = new Font("Microsoft JhengHei UI", 16F, FontStyle.Bold);
        public static readonly Font H2        = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold);
        public static readonly Font H3        = new Font("Segoe UI", 12F, FontStyle.Bold);
        public static readonly Font H4        = new Font("Segoe UI", 11F, FontStyle.Bold);

        // ── 內文 & 標籤 ──────────────────────────────────────────────────────
        public static readonly Font Body      = new Font("Segoe UI", 11F);
        public static readonly Font Label     = new Font("Segoe UI", 10F);
        public static readonly Font Caption   = new Font("Segoe UI", 9F);

        // ── 等寬（數值顯示）──────────────────────────────────────────────────
        public static readonly Font MonoXL    = new Font("Consolas", 16F, FontStyle.Bold);
        public static readonly Font MonoLg    = new Font("Consolas", 14F, FontStyle.Bold);
        public static readonly Font Mono      = new Font("Consolas", 11F);
        public static readonly Font MonoSm    = new Font("Consolas", 10F);
        public static readonly Font MonoXS    = new Font("Consolas", 9F);

        // ── 導覽 & 中文 ──────────────────────────────────────────────────────
        public static readonly Font NavItem       = new Font("Microsoft JhengHei UI", 10.5F);
        public static readonly Font NavItemActive = new Font("Microsoft JhengHei UI", 10.5F, FontStyle.Bold);
        public static readonly Font CJKBody       = new Font("Microsoft JhengHei UI", 11.5F);
        public static readonly Font CJKBold       = new Font("Microsoft JhengHei UI", 11F, FontStyle.Bold);

        // ── 特殊 ─────────────────────────────────────────────────────────────
        public static readonly Font Price     = new Font("Segoe UI", 26F, FontStyle.Bold);
        public static readonly Font PriceSm   = new Font("Segoe UI", 16F);
        public static readonly Font Badge     = new Font("Segoe UI", 8.5F, FontStyle.Bold);
    }

    /// <summary>尺寸常數</summary>
    public static class UIMetrics
    {
        public const int HeaderHeight   = 95;
        public const int NavWidth       = 230;
        public const int NavItemHeight  = 50;
        public const int NavAccentWidth = 3;

        public const int ButtonHeight   = 42;
        public const int ToolbarHeight  = 55;
        public const int StatRowHeight  = 80;
        public const int TabHeaderH     = 40;

        public const int ToastWidth     = 420;
        public const int ToastHeight    = 60;

        public const int SpacingXS      = 4;
        public const int SpacingS       = 8;
        public const int SpacingM       = 12;
        public const int SpacingL       = 20;
        public const int SpacingXL      = 28;

        public const int ChartHeight    = 340;
        public const int StatsBarHeight = 82;
    }
}
