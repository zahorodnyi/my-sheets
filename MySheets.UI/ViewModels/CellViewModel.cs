namespace MySheets.UI.ViewModels;

using CommunityToolkit.Mvvm.ComponentModel;
using MySheets.Core.Models;

public class CellViewModel : ObservableObject {
    private readonly Cell _model;

    public CellViewModel(Cell model) {
        _model = model;
    }

    public int Row => _model.Row;

    public int Col => _model.Col;

    public string Expression {
        get => _model.Expression;
        set {
            if (_model.Expression != value) {
                _model.Expression = value;
                OnPropertyChanged();
            }
        }
    }

    public object Value => _model.Value;

    public void UpdateFromModel() {
        OnPropertyChanged(nameof(Value));
    }
}