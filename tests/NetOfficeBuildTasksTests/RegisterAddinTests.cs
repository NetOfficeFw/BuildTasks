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

        private ITaskItem GetAsTaskItem(string assemblyName)
        {
            var path = Path.Combine(TestContext.CurrentContext.TestDirectory, assemblyName);
            var item = new TaskItem(path);
            return item;
        }
    }
}