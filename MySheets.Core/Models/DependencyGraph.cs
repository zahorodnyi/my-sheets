using System.Collections.Generic;
using System.Linq;

namespace MySheets.Core.Models;

public class DependencyGraph {
    private readonly Dictionary<(int, int), HashSet<(int, int)>> _dependents = new();
    private readonly Dictionary<(int, int), HashSet<(int, int)>> _dependencies = new();

    public void AddDependency(int row, int col, int depRow, int depCol) {
        if (!_dependencies.ContainsKey((row, col))) {
            _dependencies[(row, col)] = new HashSet<(int, int)>();
        }
        _dependencies[(row, col)].Add((depRow, depCol));

        if (!_dependents.ContainsKey((depRow, depCol))) {
            _dependents[(depRow, depCol)] = new HashSet<(int, int)>();
        }
        _dependents[(depRow, depCol)].Add((row, col));
    }

    public IEnumerable<(int, int)> GetDependents(int row, int col) {
        if (_dependents.TryGetValue((row, col), out var dependents)) {
            return dependents;
        }
        return Enumerable.Empty<(int, int)>();
    }

    public void ClearDependencies(int row, int col) {
        if (_dependencies.TryGetValue((row, col), out var providers)) {
            foreach (var provider in providers) {
                if (_dependents.TryGetValue(provider, out var consumers)) {
                    consumers.Remove((row, col));
                }
            }
            _dependencies.Remove((row, col));
        }
    }
}