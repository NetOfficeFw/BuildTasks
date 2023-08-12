using System;
using System.Collections;
using NUnit.Framework;

namespace NetOfficeFw.Build
{
    public class GuidExTests
    {
        [Test]
        [TestCaseSource(nameof(Guids))]
        public void ToRegistryString_ValidGuid_ReturnsGuidInWindowsRegistryFormat(Guid guid, string expectedValue)
        {
            // Arrange

            // Act
            var actualValue = guid.ToRegistryString();

            // Assert
            Assert.AreEqual(expectedValue, actualValue);
        }

        public static IEnumerable Guids()
        {
            yield return new TestCaseData(Guid.Empty, "{00000000-0000-0000-0000-000000000000}");
            yield return new TestCaseData(new Guid("62C8FE65-4EBB-45E7-B440-6E39B2CDBF29"), "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}");
            yield return new TestCaseData(new Guid("70680B7A-09A3-43EE-85AE-E21D54A1C075"), "{70680B7A-09A3-43EE-85AE-E21D54A1C075}");
        }
    }
}
