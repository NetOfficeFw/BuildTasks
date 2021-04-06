using Microsoft.Build.Utilities;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace NetOffice.Build
{
    public class ComClassRegistry
    {
        public ComClassRegistry(TaskLoggingHelper log)
        {
            this.HkeyBase = Registry.CurrentUser;
            this.Log = log;
        }

        public RegistryKey HkeyBase { get; }

        public TaskLoggingHelper Log { get; }

        public string RegisterProgId(string progId, Guid guid)
        {
            try
            {
                using var classes = this.HkeyBase.OpenSubKey(@"Software\Classes", writable: true);
                var progIdKey = classes.CreateSubKey(progId);
                progIdKey.SetValue(null, progId);

                var guidKey = progIdKey.CreateSubKey("CLSID");
                var guidValue = guid.ToRegistryString();
                guidKey.SetValue(null, guidValue);

                return progIdKey.Name;
            }
            catch (Exception ex)
            {
                Log.LogError(ex.Message);
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
                Log.LogError(ex.Message);
            }
        }

        public string RegisterComClassNative(string progId, Guid guid, Assembly assembly)
        {
            return RegisterComClass(RegistryView.Registry64, @"", progId, guid, assembly);
        }

        public string RegisterComClassWOW6432(string progId, Guid guid, Assembly assembly)
        {
            return RegisterComClass(RegistryView.Registry64, @"WOW6432Node\", progId, guid, assembly);
        }

        public string RegisterComClass(RegistryView view, string wow6432, string progId, Guid guid, Assembly assembly)
        {
            try
            {
                using var hkeyBase = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
                using var classes = this.HkeyBase.OpenSubKey($@"Software\Classes\{wow6432}CLSID", writable: true);
                var guidValue = guid.ToRegistryString();
                var clsidKey = classes.CreateSubKey(guidValue);

                clsidKey.SetValue(null, progId);
                clsidKey.CreateSubKey(@"Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}");
                var progIdKey = clsidKey.CreateSubKey("ProgId");
                progIdKey.SetValue(null, progId);

                var inprocServer = clsidKey.CreateSubKey("InprocServer32");
                inprocServer.SetValue(null, "mscoree.dll");
                inprocServer.SetValue("Assembly", assembly.GetName().FullName);
                inprocServer.SetValue("Class", progId);
                inprocServer.SetValue("Codebase", assembly.GetCodebase());
                inprocServer.SetValue("RuntimeVersion", "v4.0.30319");
                inprocServer.SetValue("ThreadingModel", "Both");

                return clsidKey.Name;
            }
            catch (Exception ex)
            {
                Log.LogError(ex.Message);
            }

            return null;
        }

        public void DeleteComClassNative(Guid guid)
        {
            DeleteComClass(RegistryView.Registry64, @"", guid);
        }

        public void DeleteComClassWOW6432(Guid guid)
        {
            DeleteComClass(RegistryView.Registry64, @"WOW6432Node\", guid);
        }

        public void DeleteComClass(RegistryView view, string wow6432, Guid guid)
        {
            try
            {
                var guidValue = guid.ToRegistryString();

                var registryPath = $@"Software\Classes\{wow6432}CLSID\{guidValue}";
                Log.LogMessage($@"Delete HKEY_CURRENT_USER\{registryPath}");

                using var hkeyBase = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
                hkeyBase.DeleteSubKeyTree(registryPath, false);
            }
            catch (Exception ex)
            {
                Log.LogError(ex.Message);
            }
        }
    }
}
