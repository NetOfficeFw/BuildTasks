using System;
using System.Reflection;
using NUnit.Framework;

namespace NetOffice.Build
{
    public class AssemblyExTests
    {
        [Test]
        public void GetCodebase_Assembly_ReturnsFileProtocolCodebase()
        {
            // Arrange
            var assembly = Assembly.GetExecutingAssembly();

            // Act
            var actualCodebase = assembly.GetCodebase();

            // Assert
            StringAssert.StartsWith("file:///", actualCodebase);
            StringAssert.DoesNotContain("\\", actualCodebase);
        }

        [Test]
        public void GetCodebase_FilePath_ReturnsFileProtocolCodebase()
        {
            // Arrange
            var path = @"c:\mypath\addin.dll";

            // Act
            var actualCodebase = path.GetCodebase();

            // Assert
            StringAssert.AreEqualIgnoringCase(@"file:///c:/mypath/addin.dll", actualCodebase);
        }
    }
}
