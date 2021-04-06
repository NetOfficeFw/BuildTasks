using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace NetOffice.Build
{
    public class CleanAddin : Task
    {
        [Required]
        public ITaskItem AssemblyPath { get; set; }

        public override bool Execute()
        {
            try
            {
                var assembly = Assembly.LoadFile(this.AssemblyPath.ItemSpec);
                var publicTypes = assembly.GetExportedTypes();

                var addinTypes = new List<TaskItem>();
                var registryWrites = new List<TaskItem>();

                foreach (var publicType in publicTypes)
                {
                    var isComVisible = IsComVisibleType(publicType);
                    if (isComVisible)
                    {
                        var guid = publicType.GUID;
                        var progId = GetProgId(publicType);

                        var guidComClass = guid.ToRegistryString();

                        Log.LogMessage($@"Cleaning {progId} {guidComClass}");


                        var comClass = new ComClassRegistry(this.Log);
                        comClass.DeleteProgId(progId);
                        comClass.DeleteComClassNative(guid);

                        if (Environment.Is64BitOperatingSystem)
                        {
                            comClass.DeleteComClassWOW6432(guid);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.LogError(ex.Message);
                return false;
            }

            return true;
        }

        private static bool IsComVisibleType(Type type)
        {
            var comVisibleAttribute = type.GetCustomAttribute<ComVisibleAttribute>();
            return comVisibleAttribute?.Value ?? false;
        }

        private static string GetProgId(Type type)
        {
            var attr = type.GetCustomAttribute<ProgIdAttribute>();
            return attr?.Value;
        }
    }
}
