using System;
using System.Drawing;

namespace PPTKalem
{
    /// <summary>
    /// Aktif çizim aracı türleri.
    /// </summary>
    public enum ToolType
    {
        Pen,
        Highlighter,
        Eraser,
        Laser
    }

    /// <summary>
    /// Çizim aracı ayarlarını merkezi olarak tutar.
    /// Ribbon, TaskPane ve DrawingEngine bu nesneyi paylaşır.
    /// </summary>
    public sealed class ToolSettings
    {
        private static readonly Lazy<ToolSettings> _instance =
            new Lazy<ToolSettings>(() => new ToolSettings());

        public static ToolSettings Instance => _instance.Value;

        private ToolSettings()
        {
            ActiveTool = ToolType.Pen;
            PenColor = Color.Red;
            PenWidth = 3f;
            Opacity = 255; // tam opak
        }

        // --- Aktif araç ---
        public ToolType ActiveTool { get; set; }

        // --- Renk ---
        private Color _penColor;
        public Color PenColor
        {
            get => _penColor;
            set
            {
                _penColor = value;
                OnSettingsChanged();
            }
        }

        // --- Kalınlık (1-20) ---
        private float _penWidth;
        public float PenWidth
        {
            get => _penWidth;
            set
            {
                _penWidth = Math.Max(1f, Math.Min(20f, value));
                OnSettingsChanged();
            }
        }

        // --- Opaklık (0-255, UI'da %10-%100 olarak gösterilir) ---
        private int _opacity;
        public int Opacity
        {
            get => _opacity;
            set
            {
                _opacity = Math.Max(25, Math.Min(255, value)); // min ~%10
                OnSettingsChanged();
            }
        }

        /// <summary>
        /// Aktif araca göre efektif renk döndürür.
        /// Fosforlu kalem → yarı-saydam sarı (alpha=100).
        /// Normal kalem → PenColor + Opacity.
        /// Silgi için kullanılmaz.
        /// </summary>
        public Color EffectiveColor
        {
            get
            {
                switch (ActiveTool)
                {
                    case ToolType.Highlighter:
                        return Color.FromArgb(100, PenColor);
                    case ToolType.Pen:
                        return Color.FromArgb(Opacity, PenColor);
                    default:
                        return Color.Transparent;
                }
            }
        }

        /// <summary>
        /// Aktif araca göre efektif kalınlık.
        /// Fosforlu kalem daha kalın çizer.
        /// </summary>
        public float EffectiveWidth
        {
            get
            {
                switch (ActiveTool)
                {
                    case ToolType.Highlighter:
                        return PenWidth * 3f;
                    case ToolType.Eraser:
                        return PenWidth * 2f;
                    default:
                        return PenWidth;
                }
            }
        }

        // --- Ayar değişim olayı ---
        public event EventHandler SettingsChanged;

        private void OnSettingsChanged()
        {
            SettingsChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
