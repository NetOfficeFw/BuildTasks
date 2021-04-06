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
            this.HkeyBase = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64);
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

                Log.LogMessage($@"Created {progIdKey.Name}");
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
            return RegisterComClass(@"", progId, guid, assembly);
        }

        public string RegisterComClassWOW6432(string progId, Guid guid, Assembly assembly)
        {
            return RegisterComClass(@"WOW6432Node\", progId, guid, assembly);
        }

        public string RegisterComClass(string wow6432, string progId, Guid guid, Assembly assembly)
        {
            try
            {
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

                Log.LogMessage($@"Created {clsidKey.Name}");
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
                Log.LogError(ex.Message);
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
                Log.LogError(ex.Message);
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
                Log.LogError(ex.Message);
            }
        }
    }
}
