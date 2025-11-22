using System.Linq;
using MySheets.Core.Models;
using Xunit;

namespace MySheets.Tests {
    public class DependencyGraphTests {
        [Fact]
        public void AddDependency_ShouldRegisterRelationship() {
            var graph = new DependencyGraph();
            graph.AddDependency(1, 1, 2, 2);

            var dependents = graph.GetDependents(2, 2).ToList();

            Assert.Single(dependents);
            Assert.Contains((1, 1), dependents);
        }

        [Fact]
        public void GetDependents_ShouldReturnMultipleDependents() {
            var graph = new DependencyGraph();
            graph.AddDependency(1, 1, 3, 3);
            graph.AddDependency(2, 2, 3, 3);

            var dependents = graph.GetDependents(3, 3).ToList();

            Assert.Equal(2, dependents.Count);
            Assert.Contains((1, 1), dependents);
            Assert.Contains((2, 2), dependents);
        }

        [Fact]
        public void ClearDependencies_ShouldRemoveForwardAndBackwardLinks() {
            var graph = new DependencyGraph();
            graph.AddDependency(1, 1, 2, 2);

            graph.ClearDependencies(1, 1);
            var dependents = graph.GetDependents(2, 2);

            Assert.Empty(dependents);
        }

        [Fact]
        public void ClearDependencies_ShouldOnlyClearSpecificCell() {
            var graph = new DependencyGraph();
            graph.AddDependency(1, 1, 5, 5);
            graph.AddDependency(2, 2, 5, 5);

            graph.ClearDependencies(1, 1);
            var dependents = graph.GetDependents(5, 5).ToList();

            Assert.Single(dependents);
            Assert.Contains((2, 2), dependents);
        }
    }
}