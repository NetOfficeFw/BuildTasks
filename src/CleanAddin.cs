using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace NetOffice.Build
{
    public class CleanAddin : Task
    {
        [Required]
        public ITaskItem AssemblyPath { get; set; }

        public ITaskItem[] OfficeApps { get; set; }

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
                    var isComVisible = publicType.IsComVisibleType();
                    var isAddinType = publicType.IsComAddinType();

                    if (isComVisible)
                    {
                        var guid = publicType.GUID;
                        var progId = publicType.GetProgId();

                        var guidComClass = guid.ToRegistryString();

                        Log.LogMessage(MessageImportance.High, $@"Cleaning {progId} {guidComClass}");

                        var comClass = new ComClassRegistry(this.Log);
                        comClass.DeleteProgId(progId);
                        comClass.DeleteComClassNative(guid);

                        if (Environment.Is64BitOperatingSystem)
                        {
                            comClass.DeleteComClassWOW6432(guid);
                        }

                        if (isAddinType && this.OfficeApps != null)
                        {
                            foreach (var officeAppItem in this.OfficeApps)
                            {
                                var officeApp = officeAppItem.ItemSpec;
                                comClass.DeleteOfficeAddin(officeApp, progId);
                            }
                        }
                    }
                }
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
