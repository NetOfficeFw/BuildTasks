using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace NetOfficeBuildTasks
{
    public class RegisterAddin : Task
    {
        [Required]
        public ITaskItem AssemblyPath { get; set; }

        [Output]
        public ITaskItem[] AddinTypes { get; set; }

        [Output]
        public ITaskItem[] RegistryWrites { get; set; }

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
                        var name = publicType.Name;
                        var guid = publicType.GUID;
                        var progId = GetProgId(publicType);

                        var metadata = new Dictionary<string,string>
                        {
                            { "Namespace", publicType.Namespace },
                            { "Guid", guid.ToString("D") },
                            { "ProgId", progId }
                        };

                        var itemWrite = new TaskItem(name, metadata);
                        addinTypes.Add(itemWrite);

                        var comClass = new ComClassRegistry(this.Log);
                        var registeredKey = comClass.RegisterProgId(progId, guid);

                        if (registeredKey != null)
                        {
                            var registryWrite = new TaskItem(registeredKey);
                            registryWrites.Add(registryWrite);
                        }
                    }
                }

                this.AddinTypes = addinTypes.ToArray();
                this.RegistryWrites = registryWrites.ToArray();
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
