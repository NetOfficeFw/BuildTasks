using System;
using System.IO;
using System.Linq;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using NUnit.Framework;

namespace NetOffice.Build
{
    public class RegisterAddinTests
    {
        [Test]
        public void Execute_Test1()
        {
            // Arrange
            var task = new RegisterAddin();
            task.AssemblyPath = GetAsTaskItem("AcmeCorpAddinNet46.dll");

            // Act
            bool run = task.Execute();

            // Assert
            Assert.IsTrue(run);
        }

        [Test]
        public void AddinTypes_Test1()
        {
            // Arrange
            var task = new RegisterAddin();
            task.AssemblyPath = GetAsTaskItem("AcmeCorpAddinNet46.dll");

            // Act
            task.Execute();
            var types = task.AddinTypes;

            // Assert
            var connectClass = types[0];
            var ns = connectClass.GetMetadata("Namespace");
            var guid = connectClass.GetMetadata("Guid");
            var progId  = connectClass.GetMetadata("ProgId");

            Assert.AreEqual("ConnectClass", connectClass.ItemSpec);
            Assert.AreEqual("AcmeCorp.Net46Sample", ns);
            Assert.AreEqual("70680b7a-09a3-43ee-85ae-e21d54a1c075", guid);
            Assert.AreEqual("AcmeCorp.Net46Sample.ConnectClass", progId);
        }

        [Test]
        public void RegistryWrites_ValidComClass()
        {
            // Arrange
            var expectedRegistryPaths = new[]
            {
                @"HKEY_CURRENT_USER\Software\Classes\AcmeCorp.Net46Sample.ConnectClass",
                @"HKEY_CURRENT_USER\Software\Classes\CLSID\{70680B7A-09A3-43EE-85AE-E21D54A1C075}",
                @"HKEY_CURRENT_USER\Software\Classes\WOW6432Node\CLSID\{70680B7A-09A3-43EE-85AE-E21D54A1C075}",
            };

            var task = new RegisterAddin();
            task.AssemblyPath = GetAsTaskItem("AcmeCorpAddinNet46.dll");

            // Act
            task.Execute();
            var writes = task.RegistryWrites;

            var actualRegistryPaths = writes.Select(i => i.ItemSpec);

            // Assert
            CollectionAssert.AreEquivalent(expectedRegistryPaths, actualRegistryPaths);
        }

        private ITaskItem GetAsTaskItem(string assemblyName)
        {
            var path = Path.Combine(TestContext.CurrentContext.TestDirectory, assemblyName);
            var item = new TaskItem(path);
            return item;
        }
    }
}