using System;
using System.Collections.Generic;
using System.Linq;

namespace MySheets.Core.Common;

public interface IUndoableAction {
    void Execute();
    void Undo();
}

public class UndoRedoManager {
    private readonly LinkedList<IUndoableAction> _undoStack = new();
    private readonly LinkedList<IUndoableAction> _redoStack = new();
    private readonly int _capacity;
    private bool _isPerformingAction;

    private CompositeAction? _currentGroup;

    public event EventHandler? StateChanged;

    public UndoRedoManager(int capacity = 10) {
        _capacity = capacity;
    }

    public bool CanUndo => _undoStack.Count > 0;
    public bool CanRedo => _redoStack.Count > 0;

    public void StartGroup() {
        if (_currentGroup == null) {
            _currentGroup = new CompositeAction();
        }
    }

    public void EndGroup() {
        if (_currentGroup != null) {
            var group = _currentGroup;
            _currentGroup = null;
            
            if (group.HasActions) {
                Execute(group);
            }
        }
    }

    public void Execute(IUndoableAction action) {
        if (_isPerformingAction) return;

        if (_currentGroup != null) {
            _currentGroup.Add(action);
            return;
        }

        _redoStack.Clear();
        _undoStack.AddLast(action);
        
        if (_undoStack.Count > _capacity) {
            _undoStack.RemoveFirst();
        }
        
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void Undo() {
        if (_undoStack.Count == 0) return;

        _isPerformingAction = true;
        try {
            var lastNode = _undoStack.Last;
            if (lastNode == null) return; 

            var action = lastNode.Value;
            _undoStack.RemoveLast();
            action.Undo();
            _redoStack.AddLast(action);
            
            if (_redoStack.Count > _capacity) {
                _redoStack.RemoveFirst();
            }
        }
        finally {
            _isPerformingAction = false;
            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public void Redo() {
        if (_redoStack.Count == 0) return;

        _isPerformingAction = true;
        try {
            var lastNode = _redoStack.Last;
            if (lastNode == null) return; 
            
            var action = lastNode.Value;
            _redoStack.RemoveLast();
            action.Execute();
            _undoStack.AddLast(action);

            if (_undoStack.Count > _capacity) {
                _undoStack.RemoveFirst();
            }
        }
        finally {
            _isPerformingAction = false;
            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}

public class CompositeAction : IUndoableAction {
    private readonly List<IUndoableAction> _actions = new();

    public bool HasActions => _actions.Count > 0;

    public void Add(IUndoableAction action) {
        _actions.Add(action);
    }

    public void Execute() {
        foreach (var action in _actions) {
            action.Execute();
        }
    }

    public void Undo() {
        for (int i = _actions.Count - 1; i >= 0; i--) {
            _actions[i].Undo();
        }
    }
}

public class CellEditAction : IUndoableAction {
    private readonly Action<string> _updateAction;
    private readonly string _oldValue;
    private readonly string _newValue;

    public CellEditAction(string oldValue, string newValue, Action<string> updateAction) {
        _oldValue = oldValue;
        _newValue = newValue;
        _updateAction = updateAction;
    }

    public void Execute() => _updateAction(_newValue);
    public void Undo() => _updateAction(_oldValue);
}

public class CellStyleAction<T> : IUndoableAction {
    private readonly Action<T> _updateAction;
    private readonly T _oldValue;
    private readonly T _newValue;

    public CellStyleAction(T oldValue, T newValue, Action<T> updateAction) {
        _oldValue = oldValue;
        _newValue = newValue;
        _updateAction = updateAction;
    }

    public void Execute() => _updateAction(_newValue);
    public void Undo() => _updateAction(_oldValue);
}