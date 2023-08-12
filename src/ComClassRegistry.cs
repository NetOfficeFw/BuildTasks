using System;
using System.Reflection;
using Microsoft.Build.Utilities;
using Microsoft.Win32;

namespace NetOfficeFw.Build
{
    public class ComClassRegistry
    {
        public ComClassRegistry(TaskLoggingHelper log)
        {
            this.HkeyBase = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64);
            this.Log = log;
        }

        public RegistryKey HkeyBase { get; }

        public TaskLoggingHelper Log { get; }

        public string RegisterProgId(string progId, Guid guid)
        {
            try
            {
                using var classes = this.HkeyBase.CreateSubKey(@"Software\Classes", writable: true);
                var progIdKey = classes.CreateSubKey(progId);
                progIdKey.SetValue(null, progId);

                var guidKey = progIdKey.CreateSubKey("CLSID");
                var guidValue = guid.ToRegistryString();
                guidKey.SetValue(null, guidValue);

                Log.LogMessage($@"Created {progIdKey.Name}");
                return progIdKey.Name;
            }
            catch (Exception ex)
            {
                Log.LogErrorFromException(ex);
            }

            return null;
        }

        public void DeleteProgId(string progId)
        {
            try
            {
                var registryPath = $@"Software\Classes\{progId}";
                Log.LogMessage($@"Delete HKEY_CURRENT_USER\{registryPath}");
                this.HkeyBase.DeleteSubKeyTree(registryPath, false);
            }
            catch (Exception ex)
            {
                Log.LogErrorFromException(ex);
            }
        }

        public string RegisterComClassNative(string progId, Guid guid, AssemblyName assemblyName, string assemblyCodebase)
        {
            return RegisterComClass(@"", progId, guid, assemblyName, assemblyCodebase);
        }

        public string RegisterComClassWOW6432(string progId, Guid guid, AssemblyName assemblyName, string assemblyCodebase)
        {
            return RegisterComClass(@"WOW6432Node\", progId, guid, assemblyName, assemblyCodebase);
        }

        public string RegisterComClass(string wow6432, string progId, Guid guid, AssemblyName assemblyName, string assemblyCodebase)
        {
            try
            {
                using var classes = this.HkeyBase.CreateSubKey($@"Software\Classes\{wow6432}CLSID", writable: true);
                var guidValue = guid.ToRegistryString();
                var clsidKey = classes.CreateSubKey(guidValue);

                clsidKey.SetValue(null, progId);
                clsidKey.CreateSubKey(@"Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}");
                var progIdKey = clsidKey.CreateSubKey("ProgId");
                progIdKey.SetValue(null, progId);

                var inprocServer = clsidKey.CreateSubKey("InprocServer32");
                inprocServer.SetValue(null, "mscoree.dll");
                inprocServer.SetValue("Assembly", assemblyName.FullName);
                inprocServer.SetValue("Class", progId);
                inprocServer.SetValue("Codebase", assemblyCodebase);
                inprocServer.SetValue("RuntimeVersion", "v4.0.30319");
                inprocServer.SetValue("ThreadingModel", "Both");

                Log.LogMessage($@"Created {clsidKey.Name}");
                return clsidKey.Name;
            }
            catch (Exception ex)
            {
                Log.LogErrorFromException(ex);
            }

            return null;
        }

        public void DeleteComClassNative(Guid guid)
        {
            DeleteComClass(@"", guid);
        }

        public void DeleteComClassWOW6432(Guid guid)
        {
            DeleteComClass(@"WOW6432Node\", guid);
        }

        public void DeleteComClass(string wow6432, Guid guid)
        {
            try
            {
                var guidValue = guid.ToRegistryString();

                var registryPath = $@"Software\Classes\{wow6432}CLSID\{guidValue}";
                Log.LogMessage($@"Delete HKEY_CURRENT_USER\{registryPath}");

                this.HkeyBase.DeleteSubKeyTree(registryPath, false);
            }
            catch (Exception ex)
            {
                Log.LogErrorFromException(ex);
            }
        }

        public string RegisterOfficeAddin(string officeApp, string progId)
        {
            try
            {
                using var addinKey = this.HkeyBase.CreateSubKey($@"Software\Microsoft\Office\{officeApp}\Addins\{progId}", writable: true);

                addinKey.SetValue("FriendlyName", progId);
                addinKey.SetValue("Description", progId);
                addinKey.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);

                Log.LogMessage($@"Created {addinKey.Name}");
                return addinKey.Name;
            }
            catch (Exception ex)
            {
                Log.LogErrorFromException(ex);
            }

            return null;
        }

        public void DeleteOfficeAddin(string officeApp, string progId)
        {
            try
            {
                var registryPath = $@"Software\Microsoft\Office\{officeApp}\Addins\{progId}";
                Log.LogMessage($@"Delete HKEY_CURRENT_USER\{registryPath}");

                this.HkeyBase.DeleteSubKeyTree(registryPath, false);
            }
            catch (Exception ex)
            {
                Log.LogErrorFromException(ex);
            }
        }
    }
}
