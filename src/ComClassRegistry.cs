using Microsoft.Build.Utilities;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
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
                using (var classes = this.HkeyBase.OpenSubKey(@"Software\Classes", writable: true))
                {
                    var progIdKey = classes.CreateSubKey(progId);
                    progIdKey.SetValue(null, progId);

                    var guidKey = progIdKey.CreateSubKey("CLSID");
                    var guidValue = guid.ToRegistryString();
                    guidKey.SetValue(null, guidValue);

                    return progIdKey.Name;
                }
            }
            catch (Exception ex)
            {
                Log.LogError(ex.Message);
            }

            return null;
        }
    }
}
