namespace MVBA.Net.InterviewQuestion
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;

    using FluentAssertions;

    using Xunit;
    using Xunit.Abstractions;

    /// <summary>
    /// Unit tests for <see cref="Type"/>s in the System.Collections.Generic namespace
    /// </summary>
    public class GenericCollectionsTests
    {
        /// <summary>
        /// The class that writes the test output
        /// </summary>
        private readonly ITestOutputHelper testOutputHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="GenericCollectionsTests"/> class.
        /// </summary>
        /// <param name="testOutputHelper">
        /// The class that writes the test output
        /// </param>
        public GenericCollectionsTests(ITestOutputHelper testOutputHelper)
        {
            this.testOutputHelper = testOutputHelper;
        }

        /// <summary>
        /// Tests to find <see cref="Type"/>s in System.Collections.Generic namespace that implement
        /// <see cref="IEnumerable"/> and have the <see cref="ComVisibleAttribute"/>.
        /// </summary>
        [Fact]
        public void EnumerablesHaveComVisibleAttrReturnsTrue()
        {
            // Setup
            // Gets the public generic classes in the System.Collections.Generic namespace that are assignable from IEnumerable
            IEnumerable<Type> genericEnumerableCollectionTypes = from type in Assembly.GetExecutingAssembly().GetTypes()
                                                                 where type.Namespace != null 
                                                                       && type.Namespace.Equals("System.Collections.Generic", StringComparison.Ordinal) 
                                                                       && type.IsClass
                                                                       && type.IsGenericType
                                                                       && type.IsPublic
                                                                       && type.IsAssignableFrom(typeof(IEnumerable))
                                                                 select type;

            // Execute
            // Determine the public generic classes in the in the System.Collections.Generic namespace that do NOT have the ComVisibleAttribute applied
            List<Type> enumerablesWithoutComVisible = genericEnumerableCollectionTypes.Where(type => !HasComVisibleAttribute(type)).ToList();

            // Verify
            // If there are any without the ComVisible attribute, write them out to the test output
            if (enumerablesWithoutComVisible.Any())
            {
                this.testOutputHelper.WriteLine(
                    $"The following classes do not have the {nameof(ComVisibleAttribute)} applied: {string.Join(Environment.NewLine, enumerablesWithoutComVisible.Select(type => type.Name))}");
            }

            // There should be NO public generic classes in the in the System.Collections.Generic namespace that do NOT have the ComVisibleAttribute applied
            enumerablesWithoutComVisible.Should().BeEmpty();
        }

        /// <summary>
        /// Checks if the given <see cref="Type"/> has the <see cref="ComVisibleAttribute"/> applied.
        /// </summary>
        /// <param name="type">
        /// The <see cref="Type"/> to check.
        /// </param>
        /// <returns>
        /// <c>True</c> if the given <see cref="Type"/> has the <see cref="ComVisibleAttribute"/> attribute applied,
        /// otherwise <c>false</c>.
        /// </returns>
        private static bool HasComVisibleAttribute(Type type)
        {
            return type.IsDefined(typeof(ComVisibleAttribute), true);
        }
    }
}