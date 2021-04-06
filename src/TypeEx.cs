using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace NetOffice.Build
{
    public static class TypeEx
    {
        public static bool IsComVisibleType(this Type type)
        {
            var comVisibleAttribute = type.GetCustomAttribute<ComVisibleAttribute>();
            return comVisibleAttribute?.Value ?? false;
        }

        public static bool IsComAddinType(this Type type)
        {
            var interfaces = type.GetInterfaces();
            return interfaces.Any(t => t.Name == "IDTExtensibility2");
        }

        public static string GetProgId(this Type type)
        {
            var attr = type.GetCustomAttribute<ProgIdAttribute>();
            return attr?.Value;
        }
    }
}
