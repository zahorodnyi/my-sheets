using System.Collections.Generic;
using System.Linq;
using MySheets.Core.Models;
using Xunit;

namespace MySheets.Tests {
    public class DependencyGraphTests {
        [Fact]
        public void TryAddDependency_ShouldRegisterRelationship() {
            var graph = new DependencyGraph();
            var result = graph.TryAddDependency(1, 1, 2, 2, out _);

            Assert.True(result);
            var dependents = graph.GetDependents(2, 2).ToList();

            Assert.Single(dependents);
            Assert.Contains((1, 1), dependents);
        }

        [Fact]
        public void GetDependents_ShouldReturnMultipleDependents() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 3, 3, out _);
            graph.TryAddDependency(2, 2, 3, 3, out _);

            var dependents = graph.GetDependents(3, 3).ToList();

            Assert.Equal(2, dependents.Count);
            Assert.Contains((1, 1), dependents);
            Assert.Contains((2, 2), dependents);
        }

        [Fact]
        public void ClearDependencies_ShouldRemoveForwardAndBackwardLinks() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);

            graph.ClearDependencies(1, 1);
            var dependents = graph.GetDependents(2, 2);

            Assert.Empty(dependents);
        }

        [Fact]
        public void ClearDependencies_ShouldOnlyClearSpecificCell() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 5, 5, out _);
            graph.TryAddDependency(2, 2, 5, 5, out _);

            graph.ClearDependencies(1, 1);
            var dependents = graph.GetDependents(5, 5).ToList();

            Assert.Single(dependents);
            Assert.Contains((2, 2), dependents);
        }
        
        [Fact]
        public void TryAddDependency_ShouldDetectCycle() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            graph.TryAddDependency(2, 2, 3, 3, out _);
            
            var result = graph.TryAddDependency(3, 3, 1, 1, out var cyclePath);

            Assert.False(result);
            Assert.NotEmpty(cyclePath);
        }
    }
}