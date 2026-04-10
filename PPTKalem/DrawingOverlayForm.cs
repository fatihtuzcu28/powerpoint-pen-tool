using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PPTKalem
{
    /// <summary>
    /// Transparent overlay form — PowerPoint penceresinin üstünde çizim yüzeyi.
    /// Per-pixel alpha ile tam saydam arka plan, çizimler opak.
    /// Mouse event'leri yakalar, çizim olmayan pikseller tıklanabilir değil.
    /// </summary>
    public class DrawingOverlayForm : Form
    {
        private readonly DrawingEngine _engine;
        private bool _isDrawing;
        private Bitmap _canvas;

        public bool LaserMode { get; set; }
        private Point _laserPoint;

        private bool _waitForLift;
        public void NotifyToolChanged() { _waitForLift = true; }

        // Win32 sabitleri
        private const int WS_EX_LAYERED = 0x80000;
        private const int WS_EX_TOPMOST = 0x8;
        private const int WS_EX_TOOLWINDOW = 0x80;
        private const int WS_EX_TRANSPARENT = 0x20;
        private const byte AC_SRC_OVER = 0x00;
        private const byte AC_SRC_ALPHA = 0x01;
        private const int ULW_ALPHA = 0x02;

        [DllImport("user32.dll")]
        private static extern bool UpdateLayeredWindow(IntPtr hwnd,
            IntPtr hdcDst, ref POINT pptDst, ref SIZE pSize,
            IntPtr hdcSrc, ref POINT pptSrc, uint crKey,
            ref BLENDFUNCTION pBlend, uint dwFlags);

        [DllImport("gdi32.dll")]
        private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hObj);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObj);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteDC(IntPtr hdc);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);
        private const uint GA_ROOT = 2;

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int X, Y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct SIZE { public int Width, Height; }

        [StructLayout(LayoutKind.Sequential)]
        private struct BLENDFUNCTION
        {
            public byte BlendOp, BlendFlags, SourceConstantAlpha, AlphaFormat;
        }

        public DrawingEngine Engine => _engine;

        public DrawingOverlayForm()
        {
            _engine = new DrawingEngine();
            InitializeOverlay();
        }

        private void InitializeOverlay()
        {
            this.Text = "PPTKalem Overlay";
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.StartPosition = FormStartPosition.Manual;
            this.Cursor = Cursors.Cross;
            this.KeyPreview = true;
            this.KeyDown += Overlay_KeyDown;

            // Timer: mouse polling ile çizim (layered window mouse event almaz)
            var timer = new Timer { Interval = 8 }; // ~120fps polling
            timer.Tick += PollMouse;
            timer.Start();
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOPMOST | WS_EX_TOOLWINDOW;
                return cp;
            }
        }

        // --- Mouse Polling ---
        // WS_EX_TRANSPARENT: overlay click-through, polling ile çizim.
        // WindowFromPoint overlay'ı atlar — toolbar/PPT gibi pencereleri döndürür.

        private void PollMouse(object sender, EventArgs e)
        {
            if (!this.Visible) return;

            try
            {
                if (LaserMode)
                {
                    _laserPoint = this.PointToClient(Cursor.Position);
                    UpdateOverlay();
                    return;
                }

                bool leftDown = (GetAsyncKeyState(0x01) & 0x8000) != 0;

                // Tablet: araç değişikliğinden sonra tam bir lift beklenir
                if (_waitForLift)
                {
                    if (leftDown) return;   // hâlâ basılı
                    _waitForLift = false;   // lifted → artık çizebilir
                }
                var screenPos = Cursor.Position;
                var clientPos = this.PointToClient(screenPos);

                // Overlay alanı dışındaysa yoksay
                if (!this.ClientRectangle.Contains(clientPos))
                {
                    if (_isDrawing)
                    {
                        _isDrawing = false;
                        _engine.EndStroke();
                        UpdateOverlay();
                    }
                    return;
                }

                // Fare toolbar veya başka TopMost pencere üstündeyse çizim yapma
                // (WS_EX_TRANSPARENT olduğu için WindowFromPoint overlay'ı atlar)
                var pt = new POINT { X = screenPos.X, Y = screenPos.Y };
                IntPtr hwndUnder = WindowFromPoint(pt);
                IntPtr rootUnder = GetAncestor(hwndUnder, GA_ROOT);
                // Overlay handle'ı dönmeyecek (transparent), PPT penceresi dönecek — sorun yok
                // Ama toolbar handle'ı dönerse çizim yapma
                if (hwndUnder != IntPtr.Zero && rootUnder != IntPtr.Zero)
                {
                    // Overlay'ın handle'ı dönmez (transparent), PPT dönerse çiz
                    // Toolbar veya settings formu dönerse çizme
                    try
                    {
                        if (Control.FromHandle(rootUnder) is Form f && f.TopMost && f != this)
                        {
                            if (_isDrawing)
                            {
                                _isDrawing = false;
                                _engine.EndStroke();
                                UpdateOverlay();
                            }
                            return;
                        }
                    }
                    catch { /* FromHandle başarısız olabilir, devam et */ }
                }

                if (leftDown && !_isDrawing)
                {
                    // Mouse down
                    _isDrawing = true;
                    _engine.BeginStroke(clientPos);
                    UpdateOverlay();
                }
                else if (leftDown && _isDrawing)
                {
                    // Mouse move
                    _engine.ContinueStroke(clientPos);
                    UpdateOverlay();
                }
                else if (!leftDown && _isDrawing)
                {
                    // Mouse up
                    _isDrawing = false;
                    _engine.EndStroke();
                    UpdateOverlay();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Overlay] PollMouse error: {ex.Message}");
            }
        }

        private void Overlay_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                HideOverlay();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.Z)
            {
                _engine.UndoRedo.Undo();
                UpdateOverlay();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.Y)
            {
                _engine.UndoRedo.Redo();
                UpdateOverlay();
                e.Handled = true;
            }
        }

        // --- Per-pixel alpha render via UpdateLayeredWindow ---

        private void UpdateOverlay()
        {
            if (!this.IsHandleCreated || this.Width <= 0 || this.Height <= 0) return;

            try
            {
                // Canvas oluştur/yeniden boyutlandır
                if (_canvas == null || _canvas.Width != this.Width || _canvas.Height != this.Height)
                {
                    _canvas?.Dispose();
                    _canvas = new Bitmap(this.Width, this.Height, PixelFormat.Format32bppArgb);
                }

                // Temiz canvas
                using (var g = Graphics.FromImage(_canvas))
                {
                    g.Clear(Color.FromArgb(1, 0, 0, 0));
                    if (LaserMode)
                    {
                        int r = 14;
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                        using (var fill = new SolidBrush(Color.FromArgb(210, 255, 30, 30)))
                        using (var glow = new SolidBrush(Color.FromArgb(80, 255, 100, 100)))
                        using (var border = new Pen(Color.FromArgb(200, 255, 255, 255), 1.5f))
                        {
                            g.FillEllipse(glow, _laserPoint.X - r - 4, _laserPoint.Y - r - 4, (r + 4) * 2, (r + 4) * 2);
                            g.FillEllipse(fill, _laserPoint.X - r, _laserPoint.Y - r, r * 2, r * 2);
                            g.DrawEllipse(border, _laserPoint.X - r, _laserPoint.Y - r, r * 2, r * 2);
                        }
                    }
                    else
                    {
                        _engine.Render(g);
                    }
                }

                // UpdateLayeredWindow ile per-pixel alpha
                IntPtr screenDC = GetDC(IntPtr.Zero);
                IntPtr memDC = CreateCompatibleDC(screenDC);
                IntPtr hBitmap = _canvas.GetHbitmap(Color.FromArgb(0));
                IntPtr oldBitmap = SelectObject(memDC, hBitmap);

                var size = new SIZE { Width = this.Width, Height = this.Height };
                var pointSrc = new POINT { X = 0, Y = 0 };
                var pointDst = new POINT { X = this.Left, Y = this.Top };
                var blend = new BLENDFUNCTION
                {
                    BlendOp = AC_SRC_OVER,
                    BlendFlags = 0,
                    SourceConstantAlpha = 255,
                    AlphaFormat = AC_SRC_ALPHA
                };

                UpdateLayeredWindow(this.Handle, screenDC, ref pointDst, ref size,
                    memDC, ref pointSrc, 0, ref blend, ULW_ALPHA);

                // Temizlik
                SelectObject(memDC, oldBitmap);
                DeleteObject(hBitmap);
                DeleteDC(memDC);
                ReleaseDC(IntPtr.Zero, screenDC);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Overlay] UpdateOverlay error: {ex.Message}");
            }
        }

        // --- Public API ---

        public void ShowOverlay(Rectangle bounds)
        {
            try
            {
                this.Bounds = bounds;
                if (!this.Visible)
                {
                    this.Show();
                }
                this.BringToFront();
                UpdateOverlay();
                Debug.WriteLine($"[Overlay] Shown at {bounds}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Overlay] ShowOverlay error: {ex.Message}");
            }
        }

        public void ShowFullScreen(Screen screen)
        {
            try
            {
                if (screen == null)
                    screen = Screen.PrimaryScreen;

                this.Bounds = screen.Bounds;
                if (!this.Visible)
                {
                    this.Show();
                }
                this.BringToFront();
                UpdateOverlay();
                Debug.WriteLine($"[Overlay] FullScreen on {screen.DeviceName}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Overlay] ShowFullScreen error: {ex.Message}");
            }
        }

        public void HideOverlay()
        {
            try
            {
                if (this.Visible)
                {
                    this.Hide();
                    Debug.WriteLine("[Overlay] Hidden");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Overlay] HideOverlay error: {ex.Message}");
            }
        }

        // Invalidate override — UpdateOverlay'ı tetikle
        public new void Invalidate()
        {
            UpdateOverlay();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _canvas?.Dispose();
                _canvas = null;
                this.KeyDown -= Overlay_KeyDown;
            }
            base.Dispose(disposing);
        }
    }
}
