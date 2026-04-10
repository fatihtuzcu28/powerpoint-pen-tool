using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTKalem
{
    /// <summary>
    /// PowerPoint COM Add-in giriş noktası.
    /// IDTExtensibility2 → yaşam döngüsü, IRibbonExtensibility → Ribbon UI.
    /// regasm ile kaydedilir, PPT açılınca otomatik yüklenir.
    /// </summary>
    [ComVisible(true)]
    [Guid("B7E43F1A-2C8D-4A5E-9F01-D3C6A8B72E14")]
    [ProgId("PPTKalem.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private PowerPoint.Application _app;
        private DrawingOverlayForm _overlay;
        private IRibbonUI _ribbon;
        private bool _slideShowActive;
        private Rectangle _slideShowBounds;

        /// <summary>
        /// Sunum modu aktif mi?
        /// </summary>
        public bool IsSlideShowActive => _slideShowActive;
        private FloatingToolbar _toolbar;

        // Win32
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        private const int INPUT_KEYBOARD = 1;
        private const uint KEYEVENTF_KEYUP = 0x0002;

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public int type;
            public InputUnion U;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)]
            public KEYBDINPUT ki;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        /// <summary>
        /// Tablet desteği: fare/kalem yok olana kadar bekler, sonra aksiyonu çalıştırır.
        /// </summary>
        private void ApplyAfterLift(Action action)
        {
            if ((GetAsyncKeyState(0x01) & 0x8000) == 0) { action(); return; }
            var t = new System.Windows.Forms.Timer { Interval = 16 };
            t.Tick += (s, e) =>
            {
                if ((GetAsyncKeyState(0x01) & 0x8000) != 0) return;
                t.Stop(); t.Dispose();
                action();
            };
            t.Start();
        }

        private void FocusSlideShowWindow()
        {
            try
            {
                if (!_slideShowActive || _app == null) return;
                if (_app.SlideShowWindows.Count == 0) return;
                IntPtr hwnd = new IntPtr(_app.SlideShowWindows[1].HWND);
                if (hwnd != IntPtr.Zero) SetForegroundWindow(hwnd);
            }
            catch { }
        }

        private void SendKeyToForeground(ushort vk)
        {
            try
            {
                var inputs = new INPUT[2];
                inputs[0] = new INPUT
                {
                    type = INPUT_KEYBOARD,
                    U = new InputUnion { ki = new KEYBDINPUT { wVk = vk, dwFlags = 0 } }
                };
                inputs[1] = new INPUT
                {
                    type = INPUT_KEYBOARD,
                    U = new InputUnion { ki = new KEYBDINPUT { wVk = vk, dwFlags = KEYEVENTF_KEYUP } }
                };
                SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
            }
            catch { }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        /// <summary>
        /// Overlay'a dışarıdan erişim (KalemRibbon callback'leri için).
        /// </summary>
        public DrawingOverlayForm Overlay => _overlay;

        /// <summary>
        /// PowerPoint Application nesnesine erişim.
        /// </summary>
        public PowerPoint.Application Application => _app;

        // Singleton — Ribbon callback'lerinden erişim için
        internal static Connect Instance { get; private set; }

        // ==================== IDTExtensibility2 ====================

        public void OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            try
            {
                Debug.WriteLine("[Connect] OnConnection");
                _app = (PowerPoint.Application)application;
                Instance = this;

                // Overlay oluştur (henüz gösterme)
                _overlay = new DrawingOverlayForm();

                // Epic Pen tarzı floating toolbar
                _toolbar = new FloatingToolbar();
                _toolbar.UpdateTheme(AppSettings.Instance.DarkTheme);
                // Araç seçimi: sunum modunda PPT native, düzenleme modunda overlay
                _toolbar.ToolSelected += (tool) =>
                {
                    ToolSettings.Instance.ActiveTool = tool;
                    if (!_slideShowActive) return; // sadece sunum modunda
                    // Tablet: hemen uygula; basılı ise lift sonrası bir daha uygula (garanti)
                    ApplyPptTool(tool);
                    FocusSlideShowWindow();
                    ApplyAfterLift(() => ApplyPptTool(tool));
                };

                _toolbar.PointerSelected += () =>
                {
                    if (!_slideShowActive) return;
                    try
                    {
                        if (_app.SlideShowWindows.Count > 0)
                        {
                            var view = _app.SlideShowWindows[1].View;
                            if (_overlay?.LaserMode == true) { _overlay.LaserMode = false; _overlay.HideOverlay(); }
                            ExecuteMsoSafe("PointerArrow");
                            view.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerArrow;
                        }
                    }
                    catch (Exception ex) { Debug.WriteLine($"[Connect] Pointer error: {ex.Message}"); }

                    FocusSlideShowWindow();

                    // Tablet: basılı ise lift sonrası tekrar uygula
                    ApplyAfterLift(() =>
                    {
                        try
                        {
                            if (_app.SlideShowWindows.Count > 0)
                            {
                                ExecuteMsoSafe("PointerArrow");
                                _app.SlideShowWindows[1].View.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerArrow;
                            }
                        }
                        catch (Exception ex) { Debug.WriteLine($"[Connect] Pointer(reapply) error: {ex.Message}"); }
                    });
                };

                _toolbar.UndoClicked += () =>
                {
                    if (!_slideShowActive) return; // sadece sunum modunda
                };
                _toolbar.RedoClicked += () =>
                {
                    if (!_slideShowActive) return; // sadece sunum modunda
                };
                _toolbar.ClearClicked += () =>
                {
                    if (!_slideShowActive) return; // sadece sunum modunda
                    try
                    {
                        if (_app.SlideShowWindows.Count > 0)
                        {
                            var view = _app.SlideShowWindows[1].View;
                            ClearInkNow(view);

                            // PPT bazen silinen mürekkebi bir sonraki stroke'ta yeniden render ediyor.
                            // Kısa bir gecikmeyle aynı işlemi tekrar uygulamak daha stabil.
                            var t = new System.Windows.Forms.Timer { Interval = 120 };
                            t.Tick += (s, e) =>
                            {
                                try
                                {
                                    ((System.Windows.Forms.Timer)s).Stop();
                                    ((System.Windows.Forms.Timer)s).Dispose();
                                    if (_app?.SlideShowWindows.Count > 0)
                                    {
                                        var v2 = _app.SlideShowWindows[1].View;
                                        ClearInkNow(v2);
                                        RefreshCurrentSlide(v2);
                                    }
                                }
                                catch { }
                            };
                            t.Start();
                        }
                    }
                    catch (Exception ex) { Debug.WriteLine($"[Connect] Erase error: {ex.Message}"); }

                    FocusSlideShowWindow();

                    // PPT'nin kendi kısayolu: E = Erase all ink on slide
                    SendKeyToForeground(0x45);

                    // Kısayol işleminden sonra tekrar temizle
                    try
                    {
                        if (_app.SlideShowWindows.Count > 0)
                        {
                            var v3 = _app.SlideShowWindows[1].View;
                            ClearInkNow(v3);
                            RefreshCurrentSlide(v3);
                        }
                    }
                    catch { }

                    var t2 = new System.Windows.Forms.Timer { Interval = 250 };
                    t2.Tick += (s, e) =>
                    {
                        try
                        {
                            ((System.Windows.Forms.Timer)s).Stop();
                            ((System.Windows.Forms.Timer)s).Dispose();
                            FocusSlideShowWindow();
                            SendKeyToForeground(0x45);
                            if (_app?.SlideShowWindows.Count > 0)
                            {
                                var v4 = _app.SlideShowWindows[1].View;
                                ClearInkNow(v4);
                                RefreshCurrentSlide(v4);
                            }
                        }
                        catch { }
                    };
                    t2.Start();
                };

                _toolbar.ColorSelected += (c) =>
                {
                    if (!_slideShowActive) return; // sadece sunum modunda
                    try
                    {
                        if (_app.SlideShowWindows.Count > 0)
                        {
                            var view = _app.SlideShowWindows[1].View;
                            ExecuteMsoSafe("PointerPen");
                            // PPT bazı sürümlerde sıraya duyarlı: önce Pen, sonra renk
                            if (view.PointerType != PowerPoint.PpSlideShowPointerType.ppSlideShowPointerPen)
                                view.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerPen;
                            SetSlideShowColor(view, c);
                        }
                    }
                    catch (Exception ex) { Debug.WriteLine($"[Connect] Color error: {ex.Message}"); }

                    FocusSlideShowWindow();

                    // Tablet: basılı ise lift sonrası tekrar uygula
                    ApplyAfterLift(() =>
                    {
                        try
                        {
                            if (_app.SlideShowWindows.Count > 0)
                            {
                                var view = _app.SlideShowWindows[1].View;
                                ExecuteMsoSafe("PointerPen");
                                if (view.PointerType != PowerPoint.PpSlideShowPointerType.ppSlideShowPointerPen)
                                    view.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerPen;
                                SetSlideShowColor(view, c);
                            }
                        }
                        catch (Exception ex) { Debug.WriteLine($"[Connect] Color(reapply) error: {ex.Message}"); }
                    });
                };

                _toolbar.CloseClicked += () => { _overlay?.HideOverlay(); _toolbar?.Hide(); _ribbon?.Invalidate(); };
                _toolbar.HotkeyToggled += () => _ribbon?.Invalidate();
                _toolbar.EndSlideShowClicked += () =>
                {
                    try
                    {
                        if (_app.SlideShowWindows.Count > 0)
                            _app.SlideShowWindows[1].View.Exit();
                    }
                    catch (Exception ex) { Debug.WriteLine($"[Connect] EndSlideShow error: {ex.Message}"); }
                };

                _toolbar.BlackoutClicked += () =>
                {
                    if (!_slideShowActive) return;
                    try
                    {
                        if (_app.SlideShowWindows.Count == 0) return;
                        var view = _app.SlideShowWindows[1].View;
                        // ppSlideShowBlackScreen = 3, ppSlideShowRunning = 1
                        if (view.State == PowerPoint.PpSlideShowState.ppSlideShowBlackScreen)
                            view.State = PowerPoint.PpSlideShowState.ppSlideShowRunning;
                        else
                            view.State = PowerPoint.PpSlideShowState.ppSlideShowBlackScreen;
                    }
                    catch (Exception ex) { Debug.WriteLine($"[Connect] Blackout error: {ex.Message}"); }

                    FocusSlideShowWindow();
                };

                _toolbar.WhiteoutClicked += () =>
                {
                    if (!_slideShowActive) return;
                    try
                    {
                        if (_app.SlideShowWindows.Count == 0) return;
                        var view = _app.SlideShowWindows[1].View;
                        // ppSlideShowWhiteScreen = 2, ppSlideShowRunning = 1
                        if (view.State == PowerPoint.PpSlideShowState.ppSlideShowWhiteScreen)
                            view.State = PowerPoint.PpSlideShowState.ppSlideShowRunning;
                        else
                            view.State = PowerPoint.PpSlideShowState.ppSlideShowWhiteScreen;
                    }
                    catch (Exception ex) { Debug.WriteLine($"[Connect] Whiteout error: {ex.Message}"); }

                    FocusSlideShowWindow();
                };

                // SlideShow eventleri
                _app.SlideShowBegin += App_SlideShowBegin;
                _app.SlideShowEnd += App_SlideShowEnd;

                Debug.WriteLine("[Connect] PPTKalem loaded successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] OnConnection error: {ex.Message}");
                MessageBox.Show($"PPTKalem yüklenemedi:\n{ex.Message}",
                    "PPTKalem Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void RefreshCurrentSlide(object view)
        {
            try
            {
                if (view == null) return;

                int idx = 0;
                try
                {
                    object slide = view.GetType().InvokeMember(
                        "Slide",
                        BindingFlags.GetProperty,
                        null,
                        view,
                        null);
                    if (slide != null)
                    {
                        object si = slide.GetType().InvokeMember(
                            "SlideIndex",
                            BindingFlags.GetProperty,
                            null,
                            slide,
                            null);
                        if (si != null) idx = Convert.ToInt32(si);
                    }
                }
                catch { }

                if (idx <= 0)
                {
                    try { dynamic v = view; idx = (int)v.Slide.SlideIndex; } catch { }
                }
                if (idx <= 0) return;

                // GotoSlide(idx, msoFalse)
                try
                {
                    view.GetType().InvokeMember(
                        "GotoSlide",
                        BindingFlags.InvokeMethod,
                        null,
                        view,
                        new object[] { idx, false });
                }
                catch
                {
                    try { dynamic v = view; v.GotoSlide(idx, false); } catch { }
                }
            }
            catch { }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                Debug.WriteLine("[Connect] OnDisconnection");

                // Event temizliği
                if (_app != null)
                {
                    _app.SlideShowBegin -= App_SlideShowBegin;
                    _app.SlideShowEnd -= App_SlideShowEnd;
                }

                // Toolbar güvenli dispose
                if (_toolbar != null)
                {
                    _toolbar.Dispose();
                    _toolbar = null;
                }

                // Overlay güvenli dispose
                if (_overlay != null)
                {
                    if (_overlay.Visible)
                        _overlay.HideOverlay();
                    _overlay.Dispose();
                    _overlay = null;
                }

                Instance = null;
                _app = null;

                Debug.WriteLine("[Connect] PPTKalem unloaded");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] OnDisconnection error: {ex.Message}");
            }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ==================== IRibbonExtensibility ====================

        public string GetCustomUI(string ribbonID)
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                using (var stream = asm.GetManifestResourceStream("PPTKalem.KalemRibbon.xml"))
                {
                    if (stream == null)
                    {
                        Debug.WriteLine("[Connect] KalemRibbon.xml resource not found!");
                        return string.Empty;
                    }
                    using (var reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] GetCustomUI error: {ex.Message}");
                return string.Empty;
            }
        }

        // ==================== RIBBON CALLBACKS ====================

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            Debug.WriteLine("[Connect] Ribbon loaded");
        }

        public void InvalidateRibbon()
        {
            try { _ribbon?.Invalidate(); }
            catch (Exception ex) { Debug.WriteLine($"[Connect] Invalidate error: {ex.Message}"); }
        }

        // --- Tool Toggle ---

        public void OnTogglePen(IRibbonControl control, bool pressed)
        {
            ToolSettings.Instance.ActiveTool = ToolType.Pen;
            if (pressed && (_overlay == null || !_overlay.Visible))
                ShowEditOverlay();
            _ribbon?.Invalidate();
        }

        public void OnToggleHighlighter(IRibbonControl control, bool pressed)
        {
            ToolSettings.Instance.ActiveTool = ToolType.Highlighter;
            if (pressed && (_overlay == null || !_overlay.Visible))
                ShowEditOverlay();
            _ribbon?.Invalidate();
        }

        public void OnToggleEraser(IRibbonControl control, bool pressed)
        {
            ToolSettings.Instance.ActiveTool = ToolType.Eraser;
            if (pressed && (_overlay == null || !_overlay.Visible))
                ShowEditOverlay();
            _ribbon?.Invalidate();
        }

        public bool GetPenPressed(IRibbonControl control)
        {
            return ToolSettings.Instance.ActiveTool == ToolType.Pen && IsOverlayVisible();
        }

        public bool GetHighlighterPressed(IRibbonControl control)
        {
            return ToolSettings.Instance.ActiveTool == ToolType.Highlighter && IsOverlayVisible();
        }

        public bool GetEraserPressed(IRibbonControl control)
        {
            return ToolSettings.Instance.ActiveTool == ToolType.Eraser && IsOverlayVisible();
        }

        // --- Actions ---

        public void OnClear(IRibbonControl control)
        {
            try
            {
                _overlay?.Engine.ClearAll();
                _overlay?.Invalidate();
            }
            catch (Exception ex) { Debug.WriteLine($"[Connect] OnClear error: {ex.Message}"); }
        }

        public void OnUndo(IRibbonControl control)
        {
            try
            {
                _overlay?.Engine.UndoRedo.Undo();
                _overlay?.Invalidate();
            }
            catch (Exception ex) { Debug.WriteLine($"[Connect] OnUndo error: {ex.Message}"); }
        }

        public void OnRedo(IRibbonControl control)
        {
            try
            {
                _overlay?.Engine.UndoRedo.Redo();
                _overlay?.Invalidate();
            }
            catch (Exception ex) { Debug.WriteLine($"[Connect] OnRedo error: {ex.Message}"); }
        }

        // --- Mode ---

        public void OnToggleOverlay(IRibbonControl control, bool pressed)
        {
            try
            {
                if (!_slideShowActive)
                {
                    _toolbar?.Hide();
                    _ribbon?.Invalidate();
                    return;
                }
                if (pressed)
                {
                    if (_toolbar != null && !_toolbar.Visible) _toolbar.Show();
                    _toolbar?.BringToFront();
                }
                else
                {
                    _toolbar?.Hide();
                }
                _ribbon?.Invalidate();
            }
            catch (Exception ex) { Debug.WriteLine($"[Connect] OnToggleOverlay error: {ex.Message}"); }
        }

        public bool GetOverlayPressed(IRibbonControl control)
        {
            try { return _toolbar != null && _toolbar.Visible; }
            catch { return false; }
        }

        // --- Ayarlar ---

        public void OnToggleAutoShow(IRibbonControl control, bool pressed)
        {
            AppSettings.Instance.AutoShowOnSlideShow = pressed;
            AppSettings.Instance.Save();
            _ribbon?.Invalidate();
        }

        public bool GetAutoShowPressed(IRibbonControl control)
        {
            return AppSettings.Instance.AutoShowOnSlideShow;
        }

        public void OnToggleDarkTheme(IRibbonControl control, bool pressed)
        {
            AppSettings.Instance.DarkTheme = pressed;
            AppSettings.Instance.Save();
            _toolbar?.UpdateTheme(pressed);
            _ribbon?.Invalidate();
        }

        public bool GetDarkThemePressed(IRibbonControl control)
        {
            return AppSettings.Instance.DarkTheme;
        }

        public void OnStartSlideShow(IRibbonControl control)
        {
            try
            {
                _app?.ActivePresentation?.SlideShowSettings.Run();
            }
            catch (Exception ex) { Debug.WriteLine($"[Connect] OnStartSlideShow error: {ex.Message}"); }
        }

        // --- Export ---

        public void OnEmbedToSlide(IRibbonControl control)
        {
            EmbedDrawingToSlide();
        }

        // ==================== OVERLAY YÖNETIMI ====================

        public void ShowEditOverlay()
        {
            try
            {
                if (_overlay == null) return;
                Rectangle bounds = GetPowerPointClientBounds();
                _overlay.ShowOverlay(bounds);

                // Floating toolbar göster (sol kenara hizala)
                if (_toolbar != null && !_toolbar.Visible)
                {
                    _toolbar.Location = new Point(
                        bounds.Left + 10,
                        bounds.Top + 60);
                    _toolbar.Show();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] ShowEditOverlay error: {ex.Message}");
            }
        }

        private Rectangle GetPowerPointClientBounds()
        {
            try
            {
                // PowerPoint penceresinin tamamını al, minimal offset
                IntPtr hwnd = new IntPtr(_app.HWND);
                if (GetWindowRect(hwnd, out RECT rect))
                {
                    int topOffset = 8;    // pencere çerçevesi
                    int bottomOffset = 8;
                    return new Rectangle(
                        rect.Left,
                        rect.Top + topOffset,
                        rect.Right - rect.Left,
                        rect.Bottom - rect.Top - topOffset - bottomOffset);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] GetPPTBounds error: {ex.Message}");
            }
            return Screen.PrimaryScreen.WorkingArea;
        }

        private void ExecuteMsoSafe(string controlId)
        {
            try
            {
                if (_app == null) return;
                // PPT'nin kendi sağ-tık menüsü aksiyonları
                _app.CommandBars.ExecuteMso(controlId);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] ExecuteMso('{controlId}') error: {ex.Message}");
            }
        }

        // ==================== PPT NATIVE ARAÇ KONTROLÜ ====================

        private void ApplyPptTool(ToolType tool)
        {
            try
            {
                if (_app.SlideShowWindows.Count == 0) return;
                var view = _app.SlideShowWindows[1].View;
                switch (tool)
                {
                    case ToolType.Laser:
                        // Overlay-tabanlı lazer noktası (tüm PPT sürümlerinde çalışır)
                        _overlay.LaserMode = true;
                        _overlay.ShowOverlay(_slideShowBounds);
                        _toolbar?.BringToFront();
                        return; // view.PointerType değiştirme
                    case ToolType.Pen:
                        if (_overlay?.LaserMode == true) { _overlay.LaserMode = false; _overlay.HideOverlay(); }
                        ExecuteMsoSafe("PointerPen");
                        view.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerPen;
                        SetSlideShowColor(view, ToolSettings.Instance.PenColor);
                        break;
                    case ToolType.Highlighter:
                        if (_overlay?.LaserMode == true) { _overlay.LaserMode = false; _overlay.HideOverlay(); }
                        try
                        {
                            view.PointerType = (PowerPoint.PpSlideShowPointerType)6;
                        }
                        catch
                        {
                            view.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerPen;
                        }
                        SetSlideShowColor(view, ToolSettings.Instance.PenColor);
                        break;
                    case ToolType.Eraser:
                        if (_overlay?.LaserMode == true) { _overlay.LaserMode = false; _overlay.HideOverlay(); }
                        ExecuteMsoSafe("PointerEraser");
                        view.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerEraser;
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] ApplyPptTool error: {ex.Message}");
            }
        }

        // ==================== SLIDESHOW EVENTLERI ====================

        private void App_SlideShowBegin(PowerPoint.SlideShowWindow wn)
        {
            try
            {
                Debug.WriteLine("[Connect] SlideShow started");
                _slideShowActive = true;

                // wn.Left/Width/Height is in POINTS (not pixels) — use HWND for real pixel bounds
                Rectangle bounds;
                try
                {
                    if (GetWindowRect(new IntPtr(wn.HWND), out RECT r))
                        bounds = new Rectangle(r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top);
                    else
                        bounds = Screen.PrimaryScreen.Bounds;
                }
                catch
                {
                    bounds = Screen.PrimaryScreen.Bounds;
                }

                _slideShowBounds = bounds;

                // Edit overlay'i kapat (laser overlay sunum sırasında ayrıca açılır)
                if (_overlay?.Visible == true && !(_overlay?.LaserMode == true))
                    _overlay.HideOverlay();

                // Toolbar'ı hemen göster (ayar bağlı)
                if (_toolbar != null && AppSettings.Instance.AutoShowOnSlideShow)
                {
                    _toolbar.SetSlideShowMode(true);
                    _toolbar.Location = new Point(bounds.Left + 10, bounds.Top + 60);
                    if (!_toolbar.Visible) _toolbar.Show();
                    _toolbar.BringToFront();

                    // PPT tam ekran animasyonu bittikten sonra tekrar üste al
                    var t = new System.Windows.Forms.Timer { Interval = 400 };
                    t.Tick += (s, e) =>
                    {
                        ((System.Windows.Forms.Timer)s).Stop();
                        ((System.Windows.Forms.Timer)s).Dispose();
                        _toolbar?.BringToFront();
                    };
                    t.Start();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] SlideShowBegin error: {ex.Message}");
            }
        }

        private void App_SlideShowEnd(PowerPoint.Presentation pres)
        {
            try
            {
                Debug.WriteLine("[Connect] SlideShow ended");
                _slideShowActive = false;
                if (_overlay?.LaserMode == true) { _overlay.LaserMode = false; _overlay.HideOverlay(); }
                _toolbar?.SetSlideShowMode(false);
                _toolbar?.Hide();
                _ribbon?.Invalidate();
                // PPT kendi çizimlerini kaydetme diyaloğunu yönetir (Keep/Discard)
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] SlideShowEnd error: {ex.Message}");
            }
        }

        // ==================== EXPORT ====================

        public void EmbedDrawingToSlide()
        {
            try
            {
                if (_overlay == null || _overlay.Engine.StrokeCount == 0)
                {
                    MessageBox.Show("Gömülecek çizim yok.", "PPTKalem",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                bool success = SlideExporter.EmbedToActiveSlide(
                    _app, _overlay.Engine,
                    _overlay.Width, _overlay.Height,
                    _slideShowActive);

                if (success)
                    MessageBox.Show("Çizim slayta gömüldü.", "PPTKalem",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    MessageBox.Show("Çizim gömülemedi.", "PPTKalem",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] EmbedDrawing error: {ex.Message}");
                MessageBox.Show($"Hata: {ex.Message}", "PPTKalem",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ==================== HELPERS ====================

        private static void SetSlideShowColor(object view, Color c)
        {
            try
            {
                if (view == null) return;

                // Office ColorFormat.RGB genelde BGR long (VBA RGB() gibi): RR + GG<<8 + BB<<16
                int ole = (c.R & 0xFF) | ((c.G & 0xFF) << 8) | ((c.B & 0xFF) << 16);

                // 1) Reflection (COM proxy ile en stabil yol)
                object drawColor = null;
                try
                {
                    drawColor = view.GetType().InvokeMember(
                        "DrawColor",
                        BindingFlags.GetProperty,
                        null,
                        view,
                        null);
                }
                catch { }

                if (drawColor != null)
                {
                    try
                    {
                        drawColor.GetType().InvokeMember(
                            "RGB",
                            BindingFlags.SetProperty,
                            null,
                            drawColor,
                            new object[] { ole });
                        return;
                    }
                    catch { }
                }

                // 1b) Bazı sürümlerde DrawColor çalışmaz, PointerColor gerekir
                object pointerColor = null;
                try
                {
                    pointerColor = view.GetType().InvokeMember(
                        "PointerColor",
                        BindingFlags.GetProperty,
                        null,
                        view,
                        null);
                }
                catch { }

                if (pointerColor != null)
                {
                    try
                    {
                        pointerColor.GetType().InvokeMember(
                            "RGB",
                            BindingFlags.SetProperty,
                            null,
                            pointerColor,
                            new object[] { ole });
                        return;
                    }
                    catch { }
                }

                // 2) Fallback: dynamic
                try
                {
                    dynamic v = view;
                    try
                    {
                        dynamic dc = v.DrawColor;
                        dc.RGB = ole;
                        return;
                    }
                    catch { }

                    dynamic pc = v.PointerColor;
                    pc.RGB = ole;
                }
                catch (Exception ex2)
                {
                    Debug.WriteLine($"[Connect] SetDrawColor failed: {ex2.Message}");
                }
            }
            catch (Exception ex) { Debug.WriteLine($"[Connect] SetDrawColor error: {ex.Message}"); }
        }

        private static void ClearInkNow(object view)
        {
            try
            {
                if (view == null) return;
                try
                {
                    view.GetType().InvokeMember(
                        "EraseDrawing",
                        BindingFlags.InvokeMethod,
                        null,
                        view,
                        null);
                }
                catch
                {
                    try { dynamic v = view; v.EraseDrawing(); } catch { }
                }
                DeleteInkShapesOnCurrentSlide(view);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] ClearInkNow error: {ex.Message}");
            }
        }

        private static void DeleteInkShapesOnCurrentSlide(object view)
        {
            try
            {
                if (view == null) return;

                object slide = null;
                try
                {
                    slide = view.GetType().InvokeMember(
                        "Slide",
                        BindingFlags.GetProperty,
                        null,
                        view,
                        null);
                }
                catch { }

                if (slide == null)
                {
                    try { dynamic v = view; slide = v.Slide; } catch { }
                }
                if (slide == null) return;

                object shapes = null;
                try
                {
                    shapes = slide.GetType().InvokeMember(
                        "Shapes",
                        BindingFlags.GetProperty,
                        null,
                        slide,
                        null);
                }
                catch { }
                if (shapes == null)
                {
                    try { dynamic s = slide; shapes = s.Shapes; } catch { }
                }
                if (shapes == null) return;

                int count = 0;
                try
                {
                    object c = shapes.GetType().InvokeMember(
                        "Count",
                        BindingFlags.GetProperty,
                        null,
                        shapes,
                        null);
                    if (c != null) count = Convert.ToInt32(c);
                }
                catch
                {
                    try { dynamic sh = shapes; count = (int)sh.Count; } catch { }
                }
                if (count <= 0) return;

                // msoInkComment = 19
                const int msoInkComment = 19;

                // Silerken sondan başla
                for (int i = count; i >= 1; i--)
                {
                    object shape = null;
                    try
                    {
                        shape = shapes.GetType().InvokeMember(
                            "Item",
                            BindingFlags.InvokeMethod,
                            null,
                            shapes,
                            new object[] { i });
                    }
                    catch
                    {
                        try { dynamic sh = shapes; shape = sh.Item(i); } catch { }
                    }
                    if (shape == null) continue;

                    int type = 0;
                    try
                    {
                        object t = shape.GetType().InvokeMember(
                            "Type",
                            BindingFlags.GetProperty,
                            null,
                            shape,
                            null);
                        if (t != null) type = Convert.ToInt32(t);
                    }
                    catch
                    {
                        try { dynamic shp = shape; type = (int)shp.Type; } catch { }
                    }

                    if (type != msoInkComment) continue;

                    try
                    {
                        shape.GetType().InvokeMember(
                            "Delete",
                            BindingFlags.InvokeMethod,
                            null,
                            shape,
                            null);
                    }
                    catch
                    {
                        try { dynamic shp = shape; shp.Delete(); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] DeleteInkShapes error: {ex.Message}");
            }
        }

        private bool IsOverlayVisible()
        {
            try { return _overlay != null && _overlay.Visible; }
            catch { return false; }
        }

        // ==================== COM KAYIT ====================

        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            try
            {
                string subKey = @"Software\Microsoft\Office\PowerPoint\Addins\PPTKalem.Connect";
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(subKey))
                {
                    key.SetValue("FriendlyName", "PPTKalem - Kalem Araçları");
                    key.SetValue("Description", "PowerPoint üzerinde çizim/kalem aracı");
                    key.SetValue("LoadBehavior", 3, Microsoft.Win32.RegistryValueKind.DWord); // Otomatik yükle
                }
                Debug.WriteLine("[Connect] COM registered successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] RegisterFunction error: {ex.Message}");
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                string subKey = @"Software\Microsoft\Office\PowerPoint\Addins\PPTKalem.Connect";
                Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(subKey, false);
                Debug.WriteLine("[Connect] COM unregistered successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Connect] UnregisterFunction error: {ex.Message}");
            }
        }
    }
}
