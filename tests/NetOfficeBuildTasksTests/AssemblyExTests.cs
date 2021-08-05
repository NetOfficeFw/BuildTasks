using System;
using System.Reflection;
using NUnit.Framework;

namespace NetOffice.Build
{
    public class AssemblyExTests
    {
        [Test]
        public void GetCodebase()
        {
            // Arrange
            var assembly = Assembly.GetExecutingAssembly();

            // Act
            var actualCodebase = assembly.GetCodebase();

            // Assert
            StringAssert.StartsWith("file:///", actualCodebase);
            StringAssert.DoesNotContain("\\", actualCodebase);
        }
    }
}
