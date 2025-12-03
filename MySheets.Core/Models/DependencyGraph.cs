using System.Collections.Generic;
using System.Linq;

namespace MySheets.Core.Models;

public class DependencyGraph {
    private readonly Dictionary<(int, int), HashSet<(int, int)>> _dependents = new();
    private readonly Dictionary<(int, int), HashSet<(int, int)>> _dependencies = new();

    public bool TryAddDependency(int row, int col, int depRow, int depCol, out List<(int, int)> cyclePath) {
        cyclePath = new List<(int, int)>();

        if (row == depRow && col == depCol) {
            cyclePath.Add((row, col));
            return false;
        }

        if (TryFindCyclePath(row, col, depRow, depCol, out var path)) {
            cyclePath = path;
            cyclePath.Add((depRow, depCol)); 
            return false;
        }

        if (!_dependencies.ContainsKey((row, col))) {
            _dependencies[(row, col)] = new HashSet<(int, int)>();
        }
        _dependencies[(row, col)].Add((depRow, depCol));

        if (!_dependents.ContainsKey((depRow, depCol))) {
            _dependents[(depRow, depCol)] = new HashSet<(int, int)>();
        }
        _dependents[(depRow, depCol)].Add((row, col));

        return true;
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

    private bool TryFindCyclePath(int targetRow, int targetCol, int startRow, int startCol, out List<(int, int)> path) {
        path = new List<(int, int)>();
        var stack = new Stack<(int, int)>();
        
        var parents = new Dictionary<(int, int), (int, int)>();

        stack.Push((startRow, startCol));
        parents[(startRow, startCol)] = (-1, -1); 

        while (stack.Count > 0) {
            var current = stack.Pop();

            if (current.Item1 == targetRow && current.Item2 == targetCol) {
                ReconstructPath(parents, (startRow, startCol), (targetRow, targetCol), path);
                return true;
            }

            if (_dependencies.TryGetValue(current, out var nextNodes)) {
                foreach (var node in nextNodes) {
                    if (!parents.ContainsKey(node)) {
                        parents[node] = current;
                        stack.Push(node);
                    }
                }
            }
        }

        return false;
    }

    private void ReconstructPath(Dictionary<(int, int), (int, int)> parents, (int, int) start, (int, int) end, List<(int, int)> path) {
        var current = end;
        while (current != start && current != (-1, -1)) {
            path.Add(current);
            current = parents[current];
        }
        path.Add(start);
    }
}