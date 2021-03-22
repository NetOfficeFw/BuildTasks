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

                foreach (var publicType in publicTypes)
                {
                    var isComVisible = IsComVisibleType(publicType);
                    if (isComVisible)
                    {
                        var name = publicType.Name;
                        var metadata = new Dictionary<string,string>
                        {
                            { "Namespace", publicType.Namespace }
                        };

                        var itemWrite = new TaskItem(name, metadata);
                        addinTypes.Add(itemWrite);
                    }
                }

                this.AddinTypes = addinTypes.ToArray();
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
    }
}
