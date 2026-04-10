# PPTKalem — PowerPoint COM Add-in Kalem Aracı
<img width="1859" height="1057" alt="pen" src="https://github.com/user-attachments/assets/9053ca1f-198e-4511-a759-5f26583a82ef" />

PowerPoint Desktop üzerinde hem düzenleme hem sunum modunda çalışan bir öğretmen çizim/kalem aracı.  
**.NET COM Add-in** olarak çalışır — `regasm` ile kaydedilir, PowerPoint açılınca otomatik yüklenir.

## Özellikler (MVP)

- **Kalem** — serbest çizim (renk, kalınlık, opaklık ayarlı)
- **Fosforlu kalem** — yarı saydam vurgulama
- **Silgi** — stroke bazlı silme
- **Temizle** — tüm çizimleri sil
- **Geri Al / İleri Al** — Ctrl+Z / Ctrl+Y veya ribbon butonları
- **Düzenleme modu** — PowerPoint penceresine hizalı overlay
- **Sunum modu** — SlideShow başlayınca tam ekran overlay
- **Slayta gömme** — çizimleri PNG olarak aktif slayta ekle
- **Ayarlar paneli** — Task Pane'de renk paleti, kalınlık ve opaklık slider'ları

## Gereksinimler

- Windows 10/11
- Visual Studio 2022 (derleme için)
- .NET Framework 4.8
- Microsoft Office / PowerPoint (Desktop)

## Derleme

1. `PPTKalem.sln` dosyasını Visual Studio 2022 ile açın
2. `Ctrl+Shift+B` veya Build → Build Solution
3. Çıktı: `PPTKalem\bin\Debug\PPTKalem.dll` (veya Release)

## Kurulum (regasm ile kayıt)

**Yönetici olarak** komut satırı açıp:

```bat
:: Kayıt (yükle)
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" PPTKalem.dll /codebase

:: Kayıt silme (kaldır)
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" PPTKalem.dll /unregister
```

Veya proje içindeki batch dosyalarını kullanın:
- `install.bat` → DLL'i `AddIns` klasörüne kopyalar + regasm kaydı
- `uninstall.bat` → regasm kaydını siler

Kayıt sonrası PowerPoint'i açın → **"Kalem Araçları"** ribbon sekmesi otomatik görünür.

## Kullanım

1. **Çizim başlatma:** Ribbon → "Kalem Araçları" → "Çizim Aç/Kapat"
2. **Araç seçme:** Kalem / Fosforlu / Silgi butonları
3. **Ayarlar:** Task Pane'den renk/kalınlık/opaklık
4. **Geri al:** Ctrl+Z veya ribbon "Geri Al" butonu
5. **Sunum modu:** "Sunum Başlat" butonu veya normal SlideShow
6. **Slayta gömme:** "Slayta Göm" butonu — çizimleri PNG olarak aktif slayta ekler
7. **Kapatma:** ESC tuşu veya "Çizim Aç/Kapat" butonunu tekrar tıklayın

## Proje Yapısı

| Dosya | Sorumluluk |
|-------|-----------|
| `Connect.cs` | COM Add-in giriş noktası (IDTExtensibility2 + IRibbonExtensibility), ribbon callback'leri, overlay/slideshow yönetimi, COM kayıt fonksiyonları |
| `KalemRibbon.xml` | Ribbon UI XML tanımı |
| `DrawingOverlayForm.cs` | Transparent overlay form, mouse input, GDI+ render |
| `DrawingEngine.cs` | Stroke yönetimi, çizim/silgi mantığı |
| `ToolSettings.cs` | Aktif araç, renk, kalınlık, opaklık (singleton) |
| `UndoRedoManager.cs` | Stroke bazlı undo/redo |
| `SlideExporter.cs` | PNG render → slayta image shape gömme |
| `KalemTaskPane.cs` | WinForms ayar paneli (renk/kalınlık/opaklık) |

## Bilinen Sınırlamalar

- Overlay pozisyonu PPT penceresi hareket ettirildiğinde otomatik güncellenmez (overlay'ı kapatıp açın)
- Çoklu monitörde sunum ekranı tespiti PowerPoint'in SlideShowWindow koordinatlarına bağlıdır
- İlk sürümde metin ekleme, şekil araçları, lazer pointer, clipboard ve vektörel export yoktur

## Sonraki Faz

- [ ] Şekil araçları (dikdörtgen, daire, ok, çizgi)
- [ ] Metin ekleme
- [ ] Lazer pointer modu
- [ ] Kopyala / yapıştır
- [ ] Vektörel (FreeForm) export
- [ ] Pencere hareket/boyut değişimi takibi

dotnet build "d:\Programlar\Yapılan Programlar\ppt kalem\PPTKalem\PPTKalem.csproj" -c Release
