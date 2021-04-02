using System;

namespace NetOffice.Build
{
    public static class GuidEx
    {
        public static string ToRegistryString(this Guid guid)
        {
            return guid.ToString("B").ToUpperInvariant();
        }
    }
}
