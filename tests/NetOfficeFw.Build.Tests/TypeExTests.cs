using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using NUnit.Framework;

namespace NetOfficeFw.Build
{
    public class TypeExTests
    {
        public const string AcmeCorpAddinNet48Assembly = "AcmeCorpAddinNet48.dll";
        public const string AcmeAddinClass = "AcmeCorp.Net48Sample.AcmeAddin";
        public const string AcmeAddinTaskpaneClass = "AcmeCorp.Net48Sample.AcmeTaskpane";

        [Test]
        public void IsComVisibleType_BasicClass_ReturnsFalse()
        {
            // Arrange
            var type = typeof(BasicClass);

            // Act
            var result = TypeEx.IsComVisibleType(type);

            // Assert
            Assert.IsFalse(result);
        }

        [Test]
        public void IsComVisibleType_ComVisibleClass_ReturnsTrue()
        {
            // Arrange
            var type = typeof(ComVisibleClass);

            // Act
            var result = TypeEx.IsComVisibleType(type);

            // Assert
            Assert.IsTrue(result);
        }

        [Test]
        public void IsComVisibleType_ReflectionOnlyComClass_ReturnsTrue()
        {
            // Arrange
            var assembly = ReflectionOnlyLoadAssembly(AcmeCorpAddinNet48Assembly);
            var type = assembly.GetType(AcmeAddinClass);

            // Act
            var result = TypeEx.IsComVisibleType(type);

            // Assert
            Assert.IsTrue(result);
        }

        [Test]
        public void IsComAddinType_BasicClass_ReturnsFalse()
        {
            // Arrange
            var type = typeof(BasicClass);

            // Act
            var result = TypeEx.IsComAddinType(type);

            // Assert
            Assert.IsFalse(result);
        }

        [Test]
        public void IsComAddinType_ComAddinClass_ReturnsTrue()
        {
            // Arrange
            var type = typeof(ComAddinClass);

            // Act
            var result = TypeEx.IsComAddinType(type);

            // Assert
            Assert.IsTrue(result);
        }

        [Test]
        public void IsComAddinType_ReflectionOnlyComClass_ReturnsTrue()
        {
            // Arrange
            var assembly = ReflectionOnlyLoadAssembly(AcmeCorpAddinNet48Assembly);
            var type = assembly.GetType(AcmeAddinClass);

            // Act
            var result = TypeEx.IsComAddinType(type);

            // Assert
            Assert.IsTrue(result);
        }

        
        [Test]
        public void GetProgId_BasicClass_ReturnsFalse()
        {
            // Arrange
            var type = typeof(BasicClass);

            // Act
            var actualProgId = TypeEx.GetProgId(type);

            // Assert
            Assert.AreEqual("NetOfficeFw.Build.BasicClass", actualProgId);
        }

        [Test]
        public void GetProgId_ProgIdClass_ReturnsTrue()
        {
            // Arrange
            var type = typeof(ProgIdClass);

            // Act
            var actualProgId = TypeEx.GetProgId(type);

            // Assert
            Assert.AreEqual("CustomProgId", actualProgId);
        }

        [Test]
        public void GetProgId_ReflectionOnlyComClassWithProgIdAttribute_ReturnsTrue()
        {
            // Arrange
            var assembly = ReflectionOnlyLoadAssembly(AcmeCorpAddinNet48Assembly);
            var type = assembly.GetType(AcmeAddinClass);

            // Act
            var actualProgId = TypeEx.GetProgId(type);

            // Assert
            Assert.AreEqual("AcmeCorp.Net48Sample.AcmeAddin", actualProgId);
        }

        [Test]
        public void GetProgId_ReflectionOnlyComClassWithoutProgIdAttribute_ReturnsTrue()
        {
            // Arrange
            var assembly = ReflectionOnlyLoadAssembly(AcmeCorpAddinNet48Assembly);
            var type = assembly.GetType(AcmeAddinTaskpaneClass);

            // Act
            var actualProgId = TypeEx.GetProgId(type);

            // Assert
            Assert.AreEqual("AcmeCorp.Net48Sample.AcmeTaskpane", actualProgId);
        }

        private Assembly ReflectionOnlyLoadAssembly(string assemblyName)
        {
            var path = Path.Combine(TestContext.CurrentContext.TestDirectory, assemblyName);
            var assembly = Assembly.ReflectionOnlyLoadFrom(path);
            return assembly;
        }
    }

    public class BasicClass
    {
    }

    [ComVisible(true)]
    public class ComVisibleClass
    {
    }

    [ProgId("CustomProgId")]
    public class ProgIdClass
    {
    }

    public class ComAddinClass : IDTExtensibility2
    {
    }
}
