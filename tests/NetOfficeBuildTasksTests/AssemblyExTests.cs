using System;
using System.IO;
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
            var assembly = LoadAssembly("AcmeCorpAddinNet46.dll");

            // Act
            var actualCodebase = assembly.GetCodebase();

            // Assert
            StringAssert.StartsWith("file:///", actualCodebase);
            StringAssert.DoesNotContain("\\", actualCodebase);
        }

        private Assembly LoadAssembly(string assemblyName)
        {
            var path = Path.Combine(TestContext.CurrentContext.TestDirectory, assemblyName);
            var assembly = Assembly.LoadFrom(path);
            return assembly;
        }
    }
}
