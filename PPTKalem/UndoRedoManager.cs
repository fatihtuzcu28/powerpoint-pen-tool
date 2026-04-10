using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace PPTKalem
{
    /// <summary>
    /// Geri alınabilir/ileri alınabilir aksiyon arayüzü.
    /// </summary>
    public interface IUndoRedoAction
    {
        void Undo();
        void Redo();
    }

    /// <summary>
    /// Stroke bazlı undo/redo yöneticisi.
    /// </summary>
    public class UndoRedoManager
    {
        private readonly Stack<IUndoRedoAction> _undoStack = new Stack<IUndoRedoAction>();
        private readonly Stack<IUndoRedoAction> _redoStack = new Stack<IUndoRedoAction>();

        public bool CanUndo => _undoStack.Count > 0;
        public bool CanRedo => _redoStack.Count > 0;

        /// <summary>
        /// Yeni aksiyon ekle. Redo stack temizlenir.
        /// </summary>
        public void PushAction(IUndoRedoAction action)
        {
            if (action == null) return;
            _undoStack.Push(action);
            _redoStack.Clear();
        }

        /// <summary>
        /// Son aksiyonu geri al.
        /// </summary>
        public void Undo()
        {
            if (!CanUndo) return;
            try
            {
                var action = _undoStack.Pop();
                action.Undo();
                _redoStack.Push(action);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[UndoRedoManager] Undo error: {ex.Message}");
            }
        }

        /// <summary>
        /// Son geri alınan aksiyonu ileri al.
        /// </summary>
        public void Redo()
        {
            if (!CanRedo) return;
            try
            {
                var action = _redoStack.Pop();
                action.Redo();
                _undoStack.Push(action);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[UndoRedoManager] Redo error: {ex.Message}");
            }
        }

        /// <summary>
        /// Tüm geçmişi temizle.
        /// </summary>
        public void Clear()
        {
            _undoStack.Clear();
            _redoStack.Clear();
        }
    }
}
