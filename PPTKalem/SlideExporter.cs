using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTKalem
{
    /// <summary>
    /// Çizimleri PNG olarak render edip aktif slayta image shape olarak gömer.
    /// </summary>
    public static class SlideExporter
    {
        /// <summary>
        /// DrawingEngine'deki çizimleri aktif slayta göm.
        /// </summary>
        public static bool EmbedToActiveSlide(
            PowerPoint.Application app,
            DrawingEngine engine,
            int overlayWidth,
            int overlayHeight,
            bool isSlideShow = false)
        {
            if (app == null || engine == null)
            {
                Debug.WriteLine("[SlideExporter] Null parameter");
                return false;
            }

            if (engine.StrokeCount == 0)
            {
                Debug.WriteLine("[SlideExporter] No strokes to export");
                return false;
            }

            string tempPath = null;

            try
            {
                // 1. Bitmap render
                using (var bmp = engine.RenderToBitmap(overlayWidth, overlayHeight))
                {
                    tempPath = Path.Combine(Path.GetTempPath(), $"PPTKalem_{Guid.NewGuid():N}.png");
                    bmp.Save(tempPath, ImageFormat.Png);
                }

                // 2. Aktif slayt (sunum modunda SlideShowWindow'dan al)
                PowerPoint.Slide slide;
                if (isSlideShow && app.SlideShowWindows.Count > 0)
                    slide = app.SlideShowWindows[1].View.Slide as PowerPoint.Slide;
                else
                    slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

                if (slide == null)
                {
                    Debug.WriteLine("[SlideExporter] No active slide");
                    return false;
                }

                // 3. Slayt boyutları
                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;

                // 4. Image shape ekle
                var shape = slide.Shapes.AddPicture(
                    FileName: tempPath,
                    LinkToFile: MsoTriState.msoFalse,
                    SaveWithDocument: MsoTriState.msoTrue,
                    Left: 0,
                    Top: 0,
                    Width: slideWidth,
                    Height: slideHeight);

                shape.Name = $"PPTKalem_Drawing_{DateTime.Now:yyyyMMdd_HHmmss}";

                Debug.WriteLine($"[SlideExporter] Embedded to slide: {shape.Name}");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideExporter] Error: {ex.Message}");
                return false;
            }
            finally
            {
                // Temp dosyayı temizle
                try
                {
                    if (tempPath != null && File.Exists(tempPath))
                        File.Delete(tempPath);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[SlideExporter] Temp cleanup error: {ex.Message}");
                }
            }
        }
    }
}
