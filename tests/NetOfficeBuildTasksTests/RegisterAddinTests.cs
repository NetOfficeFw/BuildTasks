using System;
using System.IO;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;
using NUnit.Framework;

namespace NetOfficeBuildTasks
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
            bool run = task.Execute();
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
            var task = new RegisterAddin();
            task.AssemblyPath = GetAsTaskItem("AcmeCorpAddinNet46.dll");

            // Act
            task.Execute();
            var writes = task.RegistryWrites;

            // Assert
            var registryKeyProgId = writes[0];

            Assert.AreEqual(@"HKEY_CURRENT_USER\Software\Classes\AcmeCorp.Net46Sample.ConnectClass", registryKeyProgId.ItemSpec);
        }

        private ITaskItem GetAsTaskItem(string assemblyName)
        {
            var path = Path.Combine(TestContext.CurrentContext.TestDirectory, assemblyName);
            var item = new TaskItem(path);
            return item;
        }
    }
}