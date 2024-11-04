using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using Yubico.YubiKey;

namespace YubicoAuthService
{
    public partial class Service1 : ServiceBase
    {
        private ManagementEventWatcher usbRemoveWatcher;

        [DllImport("wtsapi32.dll", SetLastError = true)]
        private static extern bool WTSLogoffSession(IntPtr hServer, int sessionId, bool bWait);

        private List<string> connectedYubikeyIds;

        public Service1()
        {
            connectedYubikeyIds = new List<string>();
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            usbRemoveWatcher = new ManagementEventWatcher();
            usbRemoveWatcher.EventArrived += new EventArrivedEventHandler(OnUsbRemoved);
            usbRemoveWatcher.Query = new WqlEventQuery("SELECT * FROM __InstanceDeletionEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_USBControllerDevice'");
            usbRemoveWatcher.Start();


            try 
            { 
                var connectedYubikeys = YubiKeyDevice.FindAll();
                if (connectedYubikeys.Count == 0)
                {
                    // Log or handle no Yubikey found at service start
                    throw new Exception("No Yubikeys found at service start.");
                }

                // Add yubikey Ids to list
                var yubikeyIds = connectedYubikeys.Select(key => key.Id.ToString()).ToList();
                connectedYubikeyIds.AddRange(yubikeyIds);

             }
            catch (DllNotFoundException dllEx)
            {
                // Handle case where the SDK DLL is not found or not loaded correctly
                EventLog.WriteEntry("Yubikey DLL not found: " + dllEx.Message, EventLogEntryType.Error);
                Stop(); // Stop the service if SDK is missing
            }
            catch (Exception ex)
            {
                // Handle other general initialization errors
                EventLog.WriteEntry("Error initializing Yubikey SDK: " + ex.Message, EventLogEntryType.Error);
                Stop(); // Stop the service on failure
            }
        }

        protected override void OnStop()
        {
            if (usbRemoveWatcher != null)
            {
                usbRemoveWatcher.Stop();
                usbRemoveWatcher.Dispose();
            }
        }

        private void OnUsbRemoved(object sender, EventArrivedEventArgs e)
        {
            try
            {
                // Get the device instance being removed
                var instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
                string deviceId = instance["Dependent"].ToString();
                
                if (connectedYubikeyIds.Any(key => deviceId.Contains(key)))  // Match your yubikey(s) to the removed USB device here
                {
                    LogOutCurrentUser();
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("UsbLogoutService", $"Error: {ex.Message}");
            }
        }

        private void LogOutCurrentUser()
        {
            // Get the current user session ID
            int sessionId = Process.GetCurrentProcess().SessionId;

            // Log off the session
            WTSLogoffSession(IntPtr.Zero, sessionId, true);
        }

        private bool IsYubikeyPresent()
        {
            // Implement a simple check using Yubikey DLL or SDK
            try
            {
                var yubikeys = YubiKeyDevice.FindAll();
                return yubikeys.Count > 0;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
