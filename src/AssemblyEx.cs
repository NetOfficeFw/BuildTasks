using System;
using System.Reflection;

namespace NetOffice.Build
{
    public static class AssemblyEx
    {
        public static string GetCodebase(this Assembly assembly)
        {
            var path = assembly.Location.Replace('\\', '/');
            return $"file:///{path}";
        }
    }
}
