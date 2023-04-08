using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Extensibility;

namespace AcmeCorp.Net48Sample
{
    [ComVisible(true)]
    [Guid("01BF0B35-5D6E-4873-84F1-D925F39BD3AC")]
    [ProgId("AcmeCorp.Net48Sample.AcmeAddin")]
    public class AcmeAddin : IDTExtensibility2
    {
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Trace.WriteLine("AcmeCorp AcmeAddin::OnConnection (.NET Framework 4.8)");
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Trace.WriteLine("AcmeCorp AcmeAddin::OnDisconnection (.NET Framework 4.8)");
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            Trace.WriteLine("AcmeCorp AcmeAddin::OnAddInsUpdate (.NET Framework 4.8)");
        }

        public void OnStartupComplete(ref Array custom)
        {
            Trace.WriteLine("AcmeCorp AcmeAddin::OnStartupComplete (.NET Framework 4.8)");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Trace.WriteLine("AcmeCorp AcmeAddin::OnBeginShutdown (.NET Framework 4.8)");
        }
    }
}
