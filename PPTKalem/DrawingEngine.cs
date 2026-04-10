using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace PPTKalem
{
    /// <summary>
    /// Tek bir çizgi darbesini (stroke) temsil eder.
    /// </summary>
    public class Stroke
    {
        public List<PointF> Points { get; set; } = new List<PointF>();
        public Color Color { get; set; }
        public float Width { get; set; }
        public ToolType Tool { get; set; }

        public Stroke(Color color, float width, ToolType tool)
        {
            Color = color;
            Width = width;
            Tool = tool;
        }
    }

    /// <summary>
    /// Çizim motoru — stroke yönetimi, çizim/silgi mantığı, GDI+ render.
    /// UI bağımsız: mouse koordinatlarını alır, Graphics nesnesine çizer.
    /// </summary>
    public class DrawingEngine
    {
        private readonly List<Stroke> _strokes = new List<Stroke>();
        private Stroke _currentStroke;
        private readonly UndoRedoManager _undoRedo = new UndoRedoManager();

        public IReadOnlyList<Stroke> Strokes => _strokes.AsReadOnly();
        public UndoRedoManager UndoRedo => _undoRedo;

        /// <summary>
        /// Yeni stroke başlat (mouse down).
        /// </summary>
        public void BeginStroke(PointF point)
        {
            try
            {
                var settings = ToolSettings.Instance;

                if (settings.ActiveTool == ToolType.Eraser)
                {
                    // Silgi: tıklanan noktaya en yakın stroke'u sil
                    EraseStrokeAt(point);
                    return;
                }

                _currentStroke = new Stroke(
                    settings.EffectiveColor,
                    settings.EffectiveWidth,
                    settings.ActiveTool);

                _currentStroke.Points.Add(point);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DrawingEngine] BeginStroke error: {ex.Message}");
            }
        }

        /// <summary>
        /// Stroke'a nokta ekle (mouse move).
        /// </summary>
        public void ContinueStroke(PointF point)
        {
            if (_currentStroke == null) return;
            _currentStroke.Points.Add(point);
        }

        /// <summary>
        /// Stroke'u bitir (mouse up).
        /// </summary>
        public void EndStroke()
        {
            try
            {
                if (_currentStroke == null) return;
                if (_currentStroke.Points.Count < 2)
                {
                    _currentStroke = null;
                    return;
                }

                _strokes.Add(_currentStroke);
                _undoRedo.PushAction(new StrokeAddAction(_strokes, _currentStroke));
                _currentStroke = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DrawingEngine] EndStroke error: {ex.Message}");
            }
        }

        /// <summary>
        /// Tüm stroke'ları ve aktif stroke'u çizer.
        /// </summary>
        public void Render(Graphics g)
        {
            if (g == null) return;

            try
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.CompositingQuality = CompositingQuality.HighQuality;

                // Tamamlanmış stroke'lar
                foreach (var stroke in _strokes)
                {
                    DrawStroke(g, stroke);
                }

                // Aktif (devam eden) stroke
                if (_currentStroke != null)
                {
                    DrawStroke(g, _currentStroke);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DrawingEngine] Render error: {ex.Message}");
            }
        }

        /// <summary>
        /// Bitmap'e render et (export için).
        /// </summary>
        public Bitmap RenderToBitmap(int width, int height)
        {
            var bmp = new Bitmap(width, height);
            bmp.SetResolution(96, 96);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                Render(g);
            }
            return bmp;
        }

        /// <summary>
        /// Tüm stroke'ları temizle.
        /// </summary>
        public void ClearAll()
        {
            _strokes.Clear();
            _undoRedo.Clear();
            _currentStroke = null;
        }

        /// <summary>
        /// Stroke sayısı.
        /// </summary>
        public int StrokeCount => _strokes.Count;

        // --- Private helpers ---

        private void DrawStroke(Graphics g, Stroke stroke)
        {
            if (stroke.Points.Count < 2) return;

            using (var pen = new Pen(stroke.Color, stroke.Width))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                pen.LineJoin = LineJoin.Round;

                var points = stroke.Points.ToArray();
                g.DrawLines(pen, points);
            }
        }

        private void EraseStrokeAt(PointF point)
        {
            const float hitRadius = 10f;
            Stroke toRemove = null;

            // En yakın stroke'u bul (son eklenen öncelikli → tersten tara)
            for (int i = _strokes.Count - 1; i >= 0; i--)
            {
                foreach (var p in _strokes[i].Points)
                {
                    float dx = p.X - point.X;
                    float dy = p.Y - point.Y;
                    if (dx * dx + dy * dy <= hitRadius * hitRadius)
                    {
                        toRemove = _strokes[i];
                        break;
                    }
                }
                if (toRemove != null) break;
            }

            if (toRemove != null)
            {
                int idx = _strokes.IndexOf(toRemove);
                _strokes.Remove(toRemove);
                _undoRedo.PushAction(new StrokeRemoveAction(_strokes, toRemove, idx));
            }
        }
    }

    // --- Undo/Redo Action'ları ---

    /// <summary>
    /// Stroke ekleme aksiyonu (geri alınabilir).
    /// </summary>
    internal class StrokeAddAction : IUndoRedoAction
    {
        private readonly List<Stroke> _strokes;
        private readonly Stroke _stroke;

        public StrokeAddAction(List<Stroke> strokes, Stroke stroke)
        {
            _strokes = strokes;
            _stroke = stroke;
        }

        public void Undo() => _strokes.Remove(_stroke);
        public void Redo() => _strokes.Add(_stroke);
    }

    /// <summary>
    /// Stroke silme aksiyonu (geri alınabilir).
    /// </summary>
    internal class StrokeRemoveAction : IUndoRedoAction
    {
        private readonly List<Stroke> _strokes;
        private readonly Stroke _stroke;
        private readonly int _index;

        public StrokeRemoveAction(List<Stroke> strokes, Stroke stroke, int index)
        {
            _strokes = strokes;
            _stroke = stroke;
            _index = index;
        }

        public void Undo()
        {
            int idx = Math.Min(_index, _strokes.Count);
            _strokes.Insert(idx, _stroke);
        }

        public void Redo() => _strokes.Remove(_stroke);
    }
}
