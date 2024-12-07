using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace NetOfficeFw.Build
{
    public static class AssemblyEx
    {
        public static string GetCodebase(this Assembly assembly)
        {
            return GetCodebase(assembly.Location);
        }

        public static string GetCodebase(this string path)
        {
            path = path.Replace('\\', '/');
            return $"file:///{path}";
        }
    }
}
