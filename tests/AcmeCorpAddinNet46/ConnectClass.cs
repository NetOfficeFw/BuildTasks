using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Extensibility;

namespace AcmeCorp.Net46Sample
{
    [ComVisible(true)]
    [Guid("70680B7A-09A3-43EE-85AE-E21D54A1C075")]
    [ProgId("AcmeCorp.Net46Sample.ConnectClass")]

    public class ConnectClass : IDTExtensibility2
    {
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Trace.WriteLine("AcmeCorp ConnectClass::OnConnection");
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Trace.WriteLine("AcmeCorp ConnectClass::OnDisconnection");
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            Trace.WriteLine("AcmeCorp ConnectClass::OnAddInsUpdate");
        }

        public void OnStartupComplete(ref Array custom)
        {
            Trace.WriteLine("AcmeCorp ConnectClass::OnStartupComplete");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Trace.WriteLine("AcmeCorp ConnectClass::OnBeginShutdown");
        }
    }
}
