using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Policy;
using Microsoft.Build.Framework;
using Microsoft.Build.Tasks;
using Microsoft.Build.Utilities;

namespace NetOfficeFw.Build
{
    public class RegisterAddin : AppDomainIsolatedTask
    {
        [Required]
        public ITaskItem AssemblyPath { get; set; }

        public ITaskItem[] OfficeApps { get; set; }

        [Output]
        public ITaskItem[] AddinTypes { get; set; }

        [Output]
        public ITaskItem[] RegistryWrites { get; set; }

        public override bool Execute()
        {
            try
            {
                var assemblyPath = this.AssemblyPath.ItemSpec;
                var assemblyDir = Path.GetDirectoryName(assemblyPath);

                var assembly = Assembly.UnsafeLoadFrom(assemblyPath);
                var assemblyName = assembly.GetName();
                var assemblyCodebase = assemblyPath.GetCodebase();
                var publicTypes = assembly.GetExportedTypes();

                var addinTypes = new List<TaskItem>();
                var registryWrites = new List<TaskItem>();

                foreach (var publicType in publicTypes)
                {
                    var isComVisible = publicType.IsComVisibleType();
                    var isAddinType = publicType.IsComAddinType();

                    if (isComVisible)
                    {
                        var name = publicType.Name;
                        var guid = publicType.GUID;
                        var progId = publicType.GetProgId();

                        var metadata = new Dictionary<string, string>
                        {
                            { "Namespace", publicType.Namespace },
                            { "Guid", guid.ToString("D") },
                            { "ProgId", progId }
                        };

                        var itemWrite = new TaskItem(name, metadata);
                        addinTypes.Add(itemWrite);

                        Log.LogMessage(MessageImportance.High, $@"Registering class {progId} with GUID {guid.ToRegistryString()}");

                        var comClass = new ComClassRegistry(this.Log);

                        var registeredKey = comClass.RegisterProgId(progId, guid);
                        if (registeredKey != null)
                        {
                            var registryWrite = new TaskItem(registeredKey);
                            registryWrites.Add(registryWrite);
                        }

                        var comClassKey = comClass.RegisterComClassNative(progId, guid, assemblyName, assemblyCodebase);
                        if (comClassKey != null)
                        {
                            var registryWrite = new TaskItem(comClassKey);
                            registryWrites.Add(registryWrite);
                        }

                        if (Environment.Is64BitOperatingSystem)
                        {
                            var comClassWOW6432Key = comClass.RegisterComClassWOW6432(progId, guid, assemblyName, assemblyCodebase);
                            if (comClassWOW6432Key != null)
                            {
                                var registryWrite = new TaskItem(comClassWOW6432Key);
                                registryWrites.Add(registryWrite);
                            }
                        }

                        if (isAddinType && this.OfficeApps != null)
                        {
                            foreach (var officeAppItem in this.OfficeApps)
                            {
                                var officeApp = officeAppItem.ItemSpec;
                                Log.LogMessage(MessageImportance.High, $@"Registering add-in {progId} to Microsoft Office {officeApp}");

                                var addinKey = comClass.RegisterOfficeAddin(officeApp, progId);
                                if (addinKey != null)
                                {
                                    var registryWrite = new TaskItem(addinKey);
                                    registryWrites.Add(registryWrite);
                                }
                            }
                        }
                    }
                }

                this.AddinTypes = addinTypes.ToArray();
                this.RegistryWrites = registryWrites.ToArray();
            }
            catch (Exception ex)
            {
                Log.LogErrorFromException(ex);
                return false;
            }

            return true;
        }
    }
}
