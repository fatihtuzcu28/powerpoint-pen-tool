# 🖊️ PowerPoint Pen Tool

A fast, lightweight and powerful **EpicPen alternative for Microsoft PowerPoint**.

Draw, highlight, and annotate slides in real-time — both in **editing mode** and **presentation mode**.

---

![banner](img/banner-2.png)

---

## 🚀 Download

👉 **[Download Latest Version](../../releases)**

> Download the `.exe` setup file and install in seconds.

---

## 🎥 Demo

![demo](img/demo.gif)

> Real-time drawing on PowerPoint slides with smooth overlay rendering.

---

## ✨ Features

- ✍️ Freehand drawing (Pen)
- 🖍️ Highlighter (transparent drawing)
- ❌ Eraser (stroke-based)
- 🔄 Undo / Redo (Ctrl+Z / Ctrl+Y)
- 🧹 Clear all drawings
- 🎯 Works in fullscreen presentation mode
- 🖥️ Works in edit mode (window-aligned overlay)
- 📌 Export drawings to slide (PNG embed)
- 🎨 Adjustable color, thickness, opacity
- ⚡ Fast and lightweight overlay engine

---

## 🎯 Use Cases

Perfect for:

- 👨‍🏫 Teachers (online & classroom lessons)
- 🎤 Presentations
- 🎓 Live explanations
- 📊 Visual storytelling
- 🧠 Concept teaching

---

## 🧠 How It Works

This tool creates a **transparent overlay layer** on top of PowerPoint.

- Drawing happens on the overlay
- PowerPoint slides remain unchanged
- Works seamlessly during slideshow mode
- Can export drawings directly into slides

---

## 📦 Installation

### 👤 For Users (Recommended)

1. Go to 👉 **[Releases](../../releases)**
2. Download `PPTKalem_Setup.exe`
3. Run installer
4. Open PowerPoint
5. You will see **"Kalem Araçları"** tab

---

### 💻 For Developers

```bash
git clone https://github.com/fatihtuzcu28/powerpoint-pen-tool.git
```

1. Open with **Visual Studio 2022**
2. Build solution (`Ctrl + Shift + B`)
3. Register DLL:

```bat
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" PPTKalem.dll /codebase
```

---

## 🕹️ Usage

1. Open PowerPoint  
2. Go to **Kalem Araçları** tab  
3. Click **Toggle Drawing**  
4. Select tool:
   - Pen / Highlighter / Eraser  
5. Adjust settings from side panel  
6. Use:
   - `Ctrl + Shift + K` → Toggle pen toolbar  
   - `Ctrl + Z` → Undo  
   - `Ctrl + Y` → Redo  
8. Click **Export to Slide** to save drawings  

---

## 🏗️ Tech Stack

- C# (.NET Framework 4.8)
- VSTO (Visual Studio Tools for Office)
- WinForms (Overlay UI)
- GDI+ (System.Drawing)

---

## 🧩 Project Structure

```
PPTKalem/
├── Connect.cs
├── KalemRibbon.xml
├── DrawingOverlayForm.cs
├── DrawingEngine.cs
├── ToolSettings.cs
├── UndoRedoManager.cs
├── SlideExporter.cs
├── KalemTaskPane.cs
```

---

## ⚠️ Limitations

- Overlay does not auto-update when PowerPoint window moves  
- Multi-monitor detection depends on PowerPoint slideshow window  
- No shape tools or text tool (yet)  

---

## 🚧 Roadmap

- [ ] Shape tools (rectangle, circle, arrow)
- [ ] Text tool
- [ ] Laser pointer mode
- [ ] Clipboard support
- [ ] Vector export
- [ ] Auto reposition overlay

---

## 🤝 Contributing

Pull requests are welcome.  
Feel free to suggest features or improvements.

---

## ⭐ Support

If you like this project, please give it a ⭐  
It helps the project grow!

---

## 🔗 Keywords

PowerPoint, VSTO, drawing tool, pen tool, presentation, EpicPen alternative
