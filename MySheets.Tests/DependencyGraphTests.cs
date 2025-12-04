using MySheets.Core.Calculation;

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

        [Fact]
        public void TryAddDependency_SelfReference_ReturnsFalse() {
            var graph = new DependencyGraph();
            var result = graph.TryAddDependency(1, 1, 1, 1, out var cyclePath);

            Assert.False(result);
            Assert.Single(cyclePath);
            Assert.Equal((1, 1), cyclePath[0]);
        }

        [Fact]
        public void TryAddDependency_DirectCycle_ReturnsFalse() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            
            var result = graph.TryAddDependency(2, 2, 1, 1, out var cyclePath);

            Assert.False(result);
            Assert.Contains((1, 1), cyclePath);
            Assert.Contains((2, 2), cyclePath);
        }

        [Fact]
        public void TryAddDependency_TriangleCycle_ReturnsFalse() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            graph.TryAddDependency(2, 2, 3, 3, out _);
            
            var result = graph.TryAddDependency(3, 3, 1, 1, out _);

            Assert.False(result);
        }

        [Fact]
        public void TryAddDependency_FourNodeSquareCycle_ReturnsFalse() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 1, 2, out _);
            graph.TryAddDependency(1, 2, 2, 2, out _);
            graph.TryAddDependency(2, 2, 2, 1, out _);

            var result = graph.TryAddDependency(2, 1, 1, 1, out _);

            Assert.False(result);
        }

        [Fact]
        public void TryAddDependency_LongChain_NoCycle_ReturnsTrue() {
            var graph = new DependencyGraph();
            for (int i = 0; i < 50; i++) {
                var result = graph.TryAddDependency(i, 0, i + 1, 0, out _);
                Assert.True(result);
            }
        }

        [Fact]
        public void TryAddDependency_LongChain_ClosingCycle_ReturnsFalse() {
            var graph = new DependencyGraph();
            int length = 20;
            for (int i = 0; i < length; i++) {
                graph.TryAddDependency(i, 0, i + 1, 0, out _);
            }

            var result = graph.TryAddDependency(length, 0, 0, 0, out _);
            
            Assert.False(result);
        }

        [Fact]
        public void TryAddDependency_DiamondShape_ReturnsTrue() {
            var graph = new DependencyGraph();
            
            Assert.True(graph.TryAddDependency(1, 1, 2, 1, out _));
            Assert.True(graph.TryAddDependency(1, 1, 2, 2, out _));
            
            Assert.True(graph.TryAddDependency(2, 1, 3, 1, out _));
            Assert.True(graph.TryAddDependency(2, 2, 3, 1, out _));

            var deps = graph.GetDependents(3, 1).ToList();
            Assert.Equal(2, deps.Count);
        }

        [Fact]
        public void TryAddDependency_YShapeMerge_ReturnsTrue() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 3, 3, out _);
            graph.TryAddDependency(2, 2, 3, 3, out _);
            
            var result = graph.TryAddDependency(3, 3, 4, 4, out _);
            Assert.True(result);
        }

        [Fact]
        public void TryAddDependency_YShapeSplit_ReturnsTrue() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            graph.TryAddDependency(1, 1, 3, 3, out _);
            
            Assert.Single(graph.GetDependents(2, 2));
            Assert.Single(graph.GetDependents(3, 3));
        }

        [Fact]
        public void TryAddDependency_DisconnectedSubgraphs_NoCycle() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            graph.TryAddDependency(2, 2, 3, 3, out _);

            graph.TryAddDependency(10, 10, 11, 11, out _);
            graph.TryAddDependency(11, 11, 12, 12, out _);
            
            var result = graph.TryAddDependency(12, 12, 10, 10, out _);
            Assert.False(result);

            Assert.True(graph.TryAddDependency(3, 3, 10, 10, out _));
        }

        [Fact]
        public void TryAddDependency_ComplexWeb_ReturnsTrue() {
            var graph = new DependencyGraph();
            for (int i = 0; i < 5; i++) {
                for (int j = 5; j < 10; j++) {
                    Assert.True(graph.TryAddDependency(i, 0, j, 0, out _));
                }
            }
            
            Assert.Equal(5, graph.GetDependents(5, 0).Count());
        }

        [Fact]
        public void TryAddDependency_AddingDuplicateDependency_DoesNotThrow() {
            var graph = new DependencyGraph();
            Assert.True(graph.TryAddDependency(1, 1, 2, 2, out _));
            Assert.True(graph.TryAddDependency(1, 1, 2, 2, out _));

            Assert.Single(graph.GetDependents(2, 2));
        }

        [Fact]
        public void ClearDependencies_BreaksCycleChain_AllowsNewLink() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            graph.TryAddDependency(2, 2, 3, 3, out _);
            
            Assert.False(graph.TryAddDependency(3, 3, 1, 1, out _));

            graph.ClearDependencies(2, 2); 

            Assert.True(graph.TryAddDependency(3, 3, 1, 1, out _));
        }

        [Fact]
        public void ClearDependencies_NonExistentCell_DoesNotThrow() {
            var graph = new DependencyGraph();
            graph.ClearDependencies(99, 99);
            Assert.Empty(graph.GetDependents(99, 99));
        }

        [Fact]
        public void TryAddDependency_LargeCoordinates_WorksCorrectly() {
            var graph = new DependencyGraph();
            Assert.True(graph.TryAddDependency(10000, 10000, 20000, 20000, out _));
            var deps = graph.GetDependents(20000, 20000);
            Assert.Contains((10000, 10000), deps);
        }

        [Fact]
        public void GetDependents_ForLeafNode_ReturnsEmpty() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            Assert.Empty(graph.GetDependents(1, 1));
        }

        [Fact]
        public void TryAddDependency_CyclePath_ContainsAllElements() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            graph.TryAddDependency(2, 2, 3, 3, out _);
            graph.TryAddDependency(3, 3, 4, 4, out _);

            graph.TryAddDependency(4, 4, 1, 1, out var path);

            Assert.Contains((1, 1), path);
            Assert.Contains((2, 2), path);
            Assert.Contains((3, 3), path);
            Assert.Contains((4, 4), path);
        }

        [Fact]
        public void ClearDependencies_RemovesDependencyFromDependenciesList() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            graph.ClearDependencies(1, 1);
            
            graph.TryAddDependency(1, 1, 3, 3, out _);
            var depsOf2 = graph.GetDependents(2, 2);
            var depsOf3 = graph.GetDependents(3, 3);
            
            Assert.Empty(depsOf2);
            Assert.Single(depsOf3);
        }

        [Fact]
        public void TryAddDependency_MultipleCycles_DetectsFirst() {
             var graph = new DependencyGraph();
             graph.TryAddDependency(1, 1, 2, 2, out _);
             graph.TryAddDependency(2, 2, 1, 1, out var path1);
             Assert.False(path1.Count == 0);

             graph.TryAddDependency(3, 3, 3, 3, out var path2);
             Assert.False(path2.Count == 0);
        }

        [Fact]
        public void ClearDependencies_DoesNotAffectOtherNodes() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            graph.TryAddDependency(3, 3, 4, 4, out _);
            
            graph.ClearDependencies(1, 1);
            
            Assert.Single(graph.GetDependents(4, 4));
            Assert.Empty(graph.GetDependents(2, 2));
        }

        [Fact]
        public void TryAddDependency_ZeroCoordinates_Works() {
            var graph = new DependencyGraph();
            Assert.True(graph.TryAddDependency(0, 0, 1, 1, out _));
            Assert.Single(graph.GetDependents(1, 1));
        }

        [Fact]
        public void TryAddDependency_ComplexCycleExample_StemAndLoop() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            graph.TryAddDependency(2, 2, 3, 3, out _);
            graph.TryAddDependency(3, 3, 4, 4, out _);
            
            graph.TryAddDependency(4, 4, 2, 2, out var path);
            
            Assert.NotEmpty(path);
        }

        [Fact]
        public void TryAddDependency_Chain_ABC_CBA_Block() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            graph.TryAddDependency(2, 2, 3, 3, out _);
            
            bool result = graph.TryAddDependency(3, 3, 1, 1, out _);
            Assert.False(result);
        }

        [Fact]
        public void TryAddDependency_ManyToOne_NoCycle() {
            var graph = new DependencyGraph();
            for(int i=0; i<100; i++) {
                Assert.True(graph.TryAddDependency(i, 0, 1000, 0, out _));
            }
            Assert.Equal(100, graph.GetDependents(1000, 0).Count());
        }

        [Fact]
        public void TryAddDependency_OneToMany_NoCycle() {
            var graph = new DependencyGraph();
            for(int i=0; i<100; i++) {
                Assert.True(graph.TryAddDependency(1000, 0, i, 0, out _));
            }
            for(int i=0; i<100; i++) {
                Assert.Single(graph.GetDependents(i, 0));
            }
        }

        [Fact]
        public void TryAddDependency_ReusingClearedNode_Works() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            graph.ClearDependencies(1, 1);
            
            bool result = graph.TryAddDependency(1, 1, 2, 2, out _);
            Assert.True(result);
            Assert.Single(graph.GetDependents(2, 2));
        }

        [Fact]
        public void GetDependents_UnrelatedNode_ReturnsEmpty() {
            var graph = new DependencyGraph();
            graph.TryAddDependency(1, 1, 2, 2, out _);
            Assert.Empty(graph.GetDependents(3, 3));
        }
    }
}