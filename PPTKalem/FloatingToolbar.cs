using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PPTKalem
{
    /// <summary>
    /// Modern floating toolbar — Fluent Design ikonları, yumuşak gölge, yuvarlak köşeler.
    /// </summary>
    public class FloatingToolbar : Form
    {
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int CS_DROPSHADOW = 0x00020000;

        private const int BTN_SIZE = 40;
        private const int PADDING = 6;
        private const int TOOLBAR_WIDTH = BTN_SIZE + PADDING * 2;
        private const int SEP_HEIGHT = 2;
        private const int COLOR_SIZE = 18;
        private const int COLOR_PAD = 3;
        private const int COLORS_PER_ROW = 2;
        private const int CORNER_RADIUS = 12;

        private readonly List<ToolButton> _toolButtons = new List<ToolButton>();
        private readonly List<Panel> _colorPanels = new List<Panel>();
        private bool _dragging;
        private Point _dragOffset;
        private int _activeToolIndex = 0; // 0 = fare/pointer (varsayılan)
        private bool _isHorizontal = false;
        private bool _isSlideShowMode = false;
        private bool _collapsed = false;

        private readonly Color[] _colors = {
            Color.Red, Color.FromArgb(0, 180, 0), Color.FromArgb(30, 120, 255), Color.White
        };

        // Ctrl+Shift+K global kısayol
        [DllImport("user32.dll")] private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        private const int  WM_HOTKEY   = 0x0312;
        private const int  HOTKEY_ID   = 0xBF01;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT   = 0x0004;
        private const uint VK_K        = 0x4B;

        public event Action HotkeyToggled;

        public event Action<ToolType> ToolSelected;
        public event Action PointerSelected;
        public event Action UndoClicked;
        public event Action RedoClicked;
        public event Action ClearClicked;
        public event Action CloseClicked;
        public event Action<Color> ColorSelected;
        public event Action EndSlideShowClicked;
        public event Action BlackoutClicked;
        public event Action WhiteoutClicked;

        public FloatingToolbar()
        {
            InitializeToolbar();
            BuildUI();
        }

        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= WS_EX_NOACTIVATE;
                cp.ClassStyle |= CS_DROPSHADOW;
                return cp;
            }
        }

        private void InitializeToolbar()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.StartPosition = FormStartPosition.Manual;
            this.BackColor = Color.FromArgb(28, 28, 30);
            this.MinimumSize = new Size(1, 1);
            this.Size = new Size(TOOLBAR_WIDTH, 100);
            this.Width = TOOLBAR_WIDTH;
            this.Location = new Point(20, 200);
            this.Cursor = Cursors.Default;
            this.DoubleBuffered = true;
            this.Opacity = 0.96;

            this.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) { _dragging = true; _dragOffset = e.Location; } };
            this.MouseMove += (s, e) => { if (_dragging) { var p = Cursor.Position; this.Location = new Point(p.X - _dragOffset.X, p.Y - _dragOffset.Y); } };
            this.MouseUp += (s, e) => _dragging = false;

            this.KeyPreview = true;
            this.KeyDown += (s, e) => { if (e.KeyCode == Keys.Escape) CloseClicked?.Invoke(); };
            this.FormClosing += (s, e) => { e.Cancel = true; this.Hide(); CloseClicked?.Invoke(); };
        }

        private void BuildUI()
        {
            this.SuspendLayout();
            _toolButtons.Clear();
            _colorPanels.Clear();
            this.Controls.Clear();

            if (_isHorizontal)
            {
                this.MaximumSize = new Size(2000, TOOLBAR_WIDTH);
                BuildHorizontal();
            }
            else
            {
                this.MaximumSize = new Size(TOOLBAR_WIDTH, 2000);
                BuildVertical();
            }

            UpdateRegion();
            UpdateSelection();
            this.ResumeLayout(false);
        }

        // ── DİKEY mod ────────────────────────────────────────────────────────────────────
        private void BuildVertical()
        {
            int y = PADDING + 6;

            AddEyeV(ref y);

            AddActionV(ref y, "rotate", "Yatay Moda Geç", Color.FromArgb(150, 150, 200), () => ToggleOrientation());

            if (_collapsed)
            {
                y += PADDING + 4;
                this.Size = new Size(TOOLBAR_WIDTH, y);
                return;
            }

            AddV(ref y, "pointer", "Fare (Seçim)", Color.FromArgb(200, 200, 200), () =>
            {
                _activeToolIndex = 0;
                PointerSelected?.Invoke();
                UpdateSelection();
            });

            if (_isSlideShowMode)
            {
                AddV(ref y, "pen",    "Kalem", Color.FromArgb(60, 150, 255),  () => DoSelectTool(ToolType.Pen, 1));
                AddV(ref y, "eraser", "Silgi", Color.FromArgb(180, 180, 180), () => DoSelectTool(ToolType.Eraser, 2));
                y += SEP_HEIGHT;
                AddActionV(ref y, "clear", "Slayt Mürekkebini Sil", Color.FromArgb(220, 80, 80), () => ClearClicked?.Invoke());
            }
            else
            {
                AddV(ref y, "pen",    "Kalem", Color.FromArgb(60, 150, 255),  () => DoSelectTool(ToolType.Pen, 1));
                AddV(ref y, "eraser", "Silgi", Color.FromArgb(180, 180, 180), () => DoSelectTool(ToolType.Eraser, 2));
                y += SEP_HEIGHT;
                AddActionV(ref y, "undo",  "Geri Al",  Color.FromArgb(160, 160, 160), () => UndoClicked?.Invoke());
                AddActionV(ref y, "redo",  "İleri Al", Color.FromArgb(160, 160, 160), () => RedoClicked?.Invoke());
                AddActionV(ref y, "clear", "Temizle",  Color.FromArgb(220, 80, 80),   () => ClearClicked?.Invoke());
            }

            AddSepV(ref y);
            AddColorGridV(ref y);
            AddSepV(ref y);

            if (_isSlideShowMode)
                AddActionV(ref y, "blackout", "Ekranı Karart / Aç", Color.FromArgb(120, 120, 120), () => BlackoutClicked?.Invoke());
            if (_isSlideShowMode)
                AddActionV(ref y, "whiteout", "Ekranı Beyazlat / Aç", Color.FromArgb(190, 190, 190), () => WhiteoutClicked?.Invoke());

            AddActionV(ref y, "close",  "Kapat",           Color.FromArgb(220, 60, 60),   () => CloseClicked?.Invoke());
            if (_isSlideShowMode)
                AddActionV(ref y, "stop_show", "Sunumu Bitir", Color.FromArgb(255, 70, 70), () => EndSlideShowClicked?.Invoke());

            y += PADDING + 4;
            this.Size = new Size(TOOLBAR_WIDTH, y);
        }

        // ── YATAY mod ────────────────────────────────────────────────────────────────────
        private void BuildHorizontal()
        {
            int x = PADDING + 6;

            AddEyeH(ref x);

            AddActionH(ref x, "rotate", "Dikey Moda Geç", Color.FromArgb(150, 150, 200), () => ToggleOrientation());

            if (_collapsed)
            {
                x += PADDING + 4;
                this.Size = new Size(x, TOOLBAR_WIDTH);
                return;
            }

            AddH(ref x, "pointer", "Fare (Seçim)", Color.FromArgb(200, 200, 200), () =>
            {
                _activeToolIndex = 0;
                PointerSelected?.Invoke();
                UpdateSelection();
            });

            if (_isSlideShowMode)
            {
                AddH(ref x, "pen",    "Kalem", Color.FromArgb(60, 150, 255),  () => DoSelectTool(ToolType.Pen, 1));
                AddH(ref x, "eraser", "Silgi", Color.FromArgb(180, 180, 180), () => DoSelectTool(ToolType.Eraser, 2));
                x += SEP_HEIGHT;
                AddActionH(ref x, "clear", "Slayt Mürekkebini Sil", Color.FromArgb(220, 80, 80), () => ClearClicked?.Invoke());
            }
            else
            {
                AddH(ref x, "pen",    "Kalem", Color.FromArgb(60, 150, 255),  () => DoSelectTool(ToolType.Pen, 1));
                AddH(ref x, "eraser", "Silgi", Color.FromArgb(180, 180, 180), () => DoSelectTool(ToolType.Eraser, 2));
                x += SEP_HEIGHT;
                AddActionH(ref x, "undo",  "Geri Al",  Color.FromArgb(160, 160, 160), () => UndoClicked?.Invoke());
                AddActionH(ref x, "redo",  "İleri Al", Color.FromArgb(160, 160, 160), () => RedoClicked?.Invoke());
                AddActionH(ref x, "clear", "Temizle",  Color.FromArgb(220, 80, 80),   () => ClearClicked?.Invoke());
            }

            AddSepH(ref x);
            AddColorGridH(ref x);
            AddSepH(ref x);

            if (_isSlideShowMode)
                AddActionH(ref x, "blackout", "Ekranı Karart / Aç", Color.FromArgb(120, 120, 120), () => BlackoutClicked?.Invoke());
            if (_isSlideShowMode)
                AddActionH(ref x, "whiteout", "Ekranı Beyazlat / Aç", Color.FromArgb(190, 190, 190), () => WhiteoutClicked?.Invoke());

            AddActionH(ref x, "close",  "Kapat",           Color.FromArgb(220, 60, 60),   () => CloseClicked?.Invoke());
            if (_isSlideShowMode)
                AddActionH(ref x, "stop_show", "Sunumu Bitir", Color.FromArgb(255, 70, 70), () => EndSlideShowClicked?.Invoke());

            x += PADDING + 4;
            this.Size = new Size(x, TOOLBAR_WIDTH);
        }

        private void DoSelectTool(ToolType tool, int index)
        {
            _activeToolIndex = index;
            ToolSettings.Instance.ActiveTool = tool;
            ToolSelected?.Invoke(tool);
            UpdateSelection();
        }

        private void UpdateSelection()
        {
            for (int i = 0; i < _toolButtons.Count; i++)
            {
                _toolButtons[i].IsSelected = (i == _activeToolIndex);
                _toolButtons[i].Invalidate();
            }
        }

        private void ToggleCollapse()
        {
            _collapsed = !_collapsed;
            BuildUI();
        }

        // Göz butonu — seçim listesine (toolButtons) girmiyor
        private void AddEyeV(ref int y)
        {
            string icon    = _collapsed ? "eye"    : "eye_off";
            string tooltip = _collapsed ? "Göster" : "Gizle";
            Color  accent  = _collapsed ? Color.FromArgb(100, 220, 100) : Color.FromArgb(160, 200, 255);
            var btn = new ToolButton
            {
                Location    = new Point(PADDING, y),
                Size        = new Size(BTN_SIZE, BTN_SIZE),
                IconId      = icon,
                AccentColor = accent,
                ClickAction = ToggleCollapse,
                IsVertical  = true
            };
            this.Controls.Add(btn);
            new ToolTip { BackColor = Color.FromArgb(45, 45, 45), ForeColor = Color.White }.SetToolTip(btn, tooltip);
            y += BTN_SIZE + 2;
        }

        private void AddEyeH(ref int x)
        {
            string icon    = _collapsed ? "eye"    : "eye_off";
            string tooltip = _collapsed ? "Göster" : "Gizle";
            Color  accent  = _collapsed ? Color.FromArgb(100, 220, 100) : Color.FromArgb(160, 200, 255);
            var btn = new ToolButton
            {
                Location    = new Point(x, PADDING),
                Size        = new Size(BTN_SIZE, BTN_SIZE),
                IconId      = icon,
                AccentColor = accent,
                ClickAction = ToggleCollapse,
                IsVertical  = false
            };
            this.Controls.Add(btn);
            new ToolTip { BackColor = Color.FromArgb(45, 45, 45), ForeColor = Color.White }.SetToolTip(btn, tooltip);
            x += BTN_SIZE + 2;
        }

        // ── Yardımcı: seçilebilir buton (pointer, pen, eraser) ─────────────────
        private void AddV(ref int y, string iconId, string tooltip, Color accent, Action onClick)
        {
            var btn = MakeBtn(iconId, accent, onClick, true);
            btn.Location = new Point(PADDING, y);
            _toolButtons.Add(btn);
            this.Controls.Add(btn);
            new ToolTip { BackColor = Color.FromArgb(45, 45, 45), ForeColor = Color.White }.SetToolTip(btn, tooltip);
            y += BTN_SIZE + 2;
        }

        private void AddH(ref int x, string iconId, string tooltip, Color accent, Action onClick)
        {
            var btn = MakeBtn(iconId, accent, onClick, false);
            btn.Location = new Point(x, PADDING);
            _toolButtons.Add(btn);
            this.Controls.Add(btn);
            new ToolTip { BackColor = Color.FromArgb(45, 45, 45), ForeColor = Color.White }.SetToolTip(btn, tooltip);
            x += BTN_SIZE + 2;
        }

        // ── Yardımcı: aksiyon buton (seçim listesine GİRMEZ) ─────────────────
        private void AddActionV(ref int y, string iconId, string tooltip, Color accent, Action onClick)
        {
            var btn = MakeBtn(iconId, accent, onClick, true);
            btn.Location = new Point(PADDING, y);
            this.Controls.Add(btn);
            new ToolTip { BackColor = Color.FromArgb(45, 45, 45), ForeColor = Color.White }.SetToolTip(btn, tooltip);
            y += BTN_SIZE + 2;
        }

        private void AddActionH(ref int x, string iconId, string tooltip, Color accent, Action onClick)
        {
            var btn = MakeBtn(iconId, accent, onClick, false);
            btn.Location = new Point(x, PADDING);
            this.Controls.Add(btn);
            new ToolTip { BackColor = Color.FromArgb(45, 45, 45), ForeColor = Color.White }.SetToolTip(btn, tooltip);
            x += BTN_SIZE + 2;
        }

        private ToolButton MakeBtn(string iconId, Color accent, Action onClick, bool vertical)
        {
            return new ToolButton
            {
                Size = new Size(BTN_SIZE, BTN_SIZE),
                IconId = iconId,
                AccentColor = accent,
                ClickAction = onClick,
                IsVertical = vertical
            };
        }

        // ── Renk grid ──────────────────────────────────────────────────────────
        private void AddColorGridV(ref int y)
        {
            int totalW = COLORS_PER_ROW * (COLOR_SIZE + COLOR_PAD) - COLOR_PAD;
            int startX = PADDING + (BTN_SIZE - totalW) / 2;

            for (int i = 0; i < _colors.Length; i++)
            {
                int row = i / COLORS_PER_ROW;
                int col = i % COLORS_PER_ROW;
                var pnl = MakeColorPanel(_colors[i],
                    startX + col * (COLOR_SIZE + COLOR_PAD),
                    y + row * (COLOR_SIZE + COLOR_PAD));
                _colorPanels.Add(pnl);
                this.Controls.Add(pnl);
            }

            int rows = (_colors.Length + COLORS_PER_ROW - 1) / COLORS_PER_ROW;
            y += rows * (COLOR_SIZE + COLOR_PAD);
        }

        private void AddColorGridH(ref int x)
        {
            int startY = PADDING + (BTN_SIZE - COLOR_SIZE) / 2;

            for (int i = 0; i < _colors.Length; i++)
            {
                var pnl = MakeColorPanel(_colors[i],
                    x + i * (COLOR_SIZE + COLOR_PAD),
                    startY);
                _colorPanels.Add(pnl);
                this.Controls.Add(pnl);
            }

            x += _colors.Length * (COLOR_SIZE + COLOR_PAD) - COLOR_PAD;
        }

        private Panel MakeColorPanel(Color color, int px, int py)
        {
            var pnl = new Panel
            {
                Location = new Point(px, py),
                Size = new Size(COLOR_SIZE, COLOR_SIZE),
                BackColor = color,
                Cursor = Cursors.Hand
            };
            pnl.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                // Yuvarlak renk kutusu
                using (var path = RoundedRect(new Rectangle(0, 0, COLOR_SIZE - 1, COLOR_SIZE - 1), 4))
                {
                    using (var brush = new SolidBrush(color))
                        g.FillPath(brush, path);
                    using (var pen = new Pen(Color.FromArgb(50, 255, 255, 255), 1))
                        g.DrawPath(pen, path);
                }
                // Seçili gösterge — beyaz iç kenar
                if (ToolSettings.Instance.PenColor.ToArgb() == color.ToArgb())
                {
                    using (var pen = new Pen(Color.White, 2))
                    using (var path = RoundedRect(new Rectangle(2, 2, COLOR_SIZE - 5, COLOR_SIZE - 5), 3))
                        g.DrawPath(pen, path);
                }
            };
            pnl.Click += (s, e) =>
            {
                ToolSettings.Instance.PenColor = color;
                ColorSelected?.Invoke(color);
                foreach (var p in _colorPanels) p.Invalidate();
            };
            return pnl;
        }

        // ── Yön değiştirme ─────────────────────────────────────────────────────
        public void SetSlideShowMode(bool isSlideShow)
        {
            if (_isSlideShowMode == isSlideShow) return;
            _isSlideShowMode = isSlideShow;
            _activeToolIndex = 0;
            _collapsed = false; // mod değişince açık başlasın
            BuildUI();
        }

        public void UpdateTheme(bool dark)
        {
            if (dark)
            {
                this.BackColor = Color.FromArgb(28, 28, 30);
                this.Opacity = 0.96;
            }
            else
            {
                this.BackColor = Color.FromArgb(245, 245, 247);
                this.Opacity = 0.97;
            }
            BuildUI();
        }

        private void AddSepV(ref int y)
        {
            int lineY = y + SEP_HEIGHT;
            var sep = new Panel
            {
                Location = new Point(PADDING + 8, lineY),
                Size = new Size(BTN_SIZE - 16, 1),
                BackColor = Color.FromArgb(50, 255, 255, 255)
            };
            this.Controls.Add(sep);
            y += SEP_HEIGHT + 5;
        }

        private void AddSepH(ref int x)
        {
            int lineX = x + SEP_HEIGHT;
            var sep = new Panel
            {
                Location = new Point(lineX, PADDING + 8),
                Size = new Size(1, BTN_SIZE - 16),
                BackColor = Color.FromArgb(50, 255, 255, 255)
            };
            this.Controls.Add(sep);
            x += SEP_HEIGHT + 5;
        }

        private void ToggleOrientation()
        {
            _isHorizontal = !_isHorizontal;
            BuildUI();
            ClampToScreen();
        }

        private void ClampToScreen()
        {
            var screen = Screen.FromControl(this).WorkingArea;
            int x = this.Left;
            int y = this.Top;
            if (x + this.Width > screen.Right) x = screen.Right - this.Width;
            if (y + this.Height > screen.Bottom) y = screen.Bottom - this.Height;
            if (x < screen.Left) x = screen.Left;
            if (y < screen.Top) y = screen.Top;
            this.Location = new Point(x, y);
        }

        private void UpdateRegion()
        {
            using (var path = RoundedRect(new Rectangle(0, 0, this.Width, this.Height), CORNER_RADIUS))
                this.Region = new Region(path);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            bool dark = this.BackColor.GetBrightness() < 0.5f;

            // İnce kenar — hafif parlak
            using (var pen = new Pen(Color.FromArgb(dark ? 40 : 25, 255, 255, 255), 1))
            using (var path = RoundedRect(new Rectangle(0, 0, this.Width - 1, this.Height - 1), CORNER_RADIUS))
                g.DrawPath(pen, path);

            // Modern grip noktaları — yöne göre
            Color dotColor = Color.FromArgb(dark ? 80 : 140, 128, 128, 128);
            using (var brush = new SolidBrush(dotColor))
            {
                if (_isHorizontal)
                {
                    int cy = this.Height / 2;
                    for (int i = -1; i <= 1; i++)
                        g.FillEllipse(brush, 3, cy + i * 5 - 1, 3, 3);
                }
                else
                {
                    int cx = this.Width / 2;
                    for (int i = -1; i <= 1; i++)
                        g.FillEllipse(brush, cx + i * 5 - 1, 3, 3, 3);
                }
            }
        }

        protected override void OnResize(EventArgs e) { base.OnResize(e); UpdateRegion(); }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, VK_K);
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            base.OnHandleDestroyed(e);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                if (this.Visible) this.Hide();
                else { this.Show(); this.BringToFront(); }
                HotkeyToggled?.Invoke();
            }
            base.WndProc(ref m);
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    /// <summary>
    /// Toolbar butonu — Segoe MDL2 Assets font ikonları ile profesyonel render.
    /// </summary>
    public class ToolButton : Control
    {
        // Segoe MDL2 Assets monokrom ikonlar
        private static readonly Dictionary<string, string> IconMap = new Dictionary<string, string>
        {
            { "pointer",     "\uE8B0" },  // Mouse
            { "pen",         "\uEE56" },  // Pen / Inking tool
            { "highlighter", "\uE7E6" },  // Highlight
            { "eraser",      "\uED60" },  // Eraser tool
            { "undo",        "\uE7A7" },  // Undo
            { "redo",        "\uE7A6" },  // Redo
            { "clear",       "\uE74D" },  // Delete
            { "laser",       "\uE7B3" },  // Laser / target
            { "camera",      "\uE722" },  // Camera
            { "rotate",      "\uE7AD" },  // Rotate
            { "close",       "\uE711" },  // ChromeClose (X)
            { "eye",         "\uE890" },  // View (eye open)
            { "eye_off",     "\uE890" },  // View (eye — toolbar expanded)
            { "blackout",    "\uE708" },  // Brightness down (dark)
            { "whiteout",    "\uE706" },  // Brightness up (light)
            { "stop_show",   "\uE7E8" },  // Stop / exit
        };

        private static readonly string[] _iconFontNames = { "Segoe Fluent Icons", "Segoe MDL2 Assets" };
        private static string _resolvedIconFont;

        private static string GetIconFont()
        {
            if (_resolvedIconFont != null) return _resolvedIconFont;
            foreach (var name in _iconFontNames)
            {
                using (var f = new Font(name, 12f, FontStyle.Regular, GraphicsUnit.Pixel))
                {
                    if (f.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    {
                        _resolvedIconFont = name;
                        return name;
                    }
                }
            }
            _resolvedIconFont = "Segoe MDL2 Assets";
            return _resolvedIconFont;
        }

        public string IconId { get; set; }
        public Color AccentColor { get; set; } = Color.White;
        public bool IsSelected { get; set; }
        public Action ClickAction { get; set; }
        public bool IsVertical { get; set; } = true;

        private bool _hover;

        public ToolButton()
        {
            this.DoubleBuffered = true;
            this.Cursor = Cursors.Hand;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

            var rect = new Rectangle(0, 0, this.Width - 1, this.Height - 1);

            bool darkTheme = true;
            try
            {
                if (this.Parent != null)
                    darkTheme = this.Parent.BackColor.GetBrightness() < 0.5f;
            }
            catch { }

            // Arka plan — modern yuvarlak köşe
            Color bg = IsSelected
                ? Color.FromArgb(50, AccentColor.R, AccentColor.G, AccentColor.B)
                : _hover
                    ? (darkTheme ? Color.FromArgb(50, 50, 52) : Color.FromArgb(228, 228, 230))
                    : Color.Transparent;

            if (bg != Color.Transparent)
            {
                using (var brush = new SolidBrush(bg))
                using (var path = RoundedRect(rect, 7))
                    g.FillPath(brush, path);
            }

            // Seçim kenar göstergesi — yöne göre
            if (IsSelected)
            {
                using (var brush = new SolidBrush(AccentColor))
                {
                    if (IsVertical)
                    {
                        using (var path = RoundedRect(new Rectangle(0, 5, 3, this.Height - 10), 1))
                            g.FillPath(brush, path);
                    }
                    else
                    {
                        using (var path = RoundedRect(new Rectangle(5, 0, this.Width - 10, 3), 1))
                            g.FillPath(brush, path);
                    }
                }
            }

            // İkon render (Segoe MDL2 Assets / Fluent Icons)
            Color iconColor = IsSelected
                ? AccentColor
                : _hover
                    ? (darkTheme ? Color.FromArgb(255, 255, 255) : Color.FromArgb(10, 10, 10))
                    : (darkTheme ? Color.FromArgb(220, 220, 225) : Color.FromArgb(30, 30, 35));

            string glyph;
            if (IconMap.TryGetValue(IconId ?? "", out glyph))
            {
                float fontSize = this.Height * 0.48f;
                using (var font = new Font(GetIconFont(), fontSize, FontStyle.Regular, GraphicsUnit.Pixel))
                {
                    TextRenderer.DrawText(g, glyph, font,
                        new Rectangle(0, 0, this.Width, this.Height),
                        iconColor,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
                }
            }
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; this.Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; this.Invalidate(); base.OnMouseLeave(e); }
        protected override void OnClick(EventArgs e) { ClickAction?.Invoke(); base.OnClick(e); }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
