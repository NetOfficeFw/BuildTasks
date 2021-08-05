using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using NUnit.Framework;

namespace NetOffice.Build
{
    public class TypeExTests
    {
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
            var assembly = ReflectionOnlyLoadAssembly("AcmeCorpAddinNet46.dll");
            var type = assembly.GetType("AcmeCorp.Net46Sample.ConnectClass");

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
            var assembly = ReflectionOnlyLoadAssembly("AcmeCorpAddinNet46.dll");
            var type = assembly.GetType("AcmeCorp.Net46Sample.ConnectClass");

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
            Assert.AreEqual("NetOffice.Build.BasicClass", actualProgId);
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
        public void GetProgId_ReflectionOnlyComClass_ReturnsTrue()
        {
            // Arrange
            var assembly = ReflectionOnlyLoadAssembly("AcmeCorpAddinNet46.dll");
            var type = assembly.GetType("AcmeCorp.Net46Sample.ConnectClass");

            // Act
            var actualProgId = TypeEx.GetProgId(type);

            // Assert
            Assert.AreEqual("AcmeCorp.Net46Sample.ConnectClass", actualProgId);
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
