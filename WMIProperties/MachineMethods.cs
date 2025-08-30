using System;
using System.Linq;
using System.Management;
using System.Security;

namespace DDAGUI.WMIProperties
{
    public class MachineMethods
    {
        /*
         * Class properties
         */

        protected (string computerName, string username, SecureString password) userCredential;
        protected ConnectionOptions credential;
        protected ManagementScope scope;
        protected ManagementObjectSearcher searcher;

        protected string scopePath;
        protected bool isLocal;

        /*
         * Machine Method
         * 1. Empty - Local
         * 2. Or Connect through WMI
         */

        public MachineMethods()
        {
            userCredential.computerName = "localhost";
            isLocal = true;
            credential = new ConnectionOptions
            {
                Impersonation = ImpersonationLevel.Impersonate,
                Authentication = AuthenticationLevel.PacketPrivacy,
                EnablePrivileges = true
            };
        }

        public MachineMethods((string computerName, string username, SecureString password) userCredential)
        {
            this.userCredential.computerName = userCredential.computerName;
            this.userCredential.username = userCredential.username;
            this.userCredential.password = userCredential.password;
            isLocal = false;
            credential = new ConnectionOptions
            {
                Username = userCredential.username,
                SecurePassword = userCredential.password,
                Impersonation = ImpersonationLevel.Impersonate,
                Authentication = AuthenticationLevel.PacketPrivacy,
                EnablePrivileges = true
            };
        }

        /*
         * Get value
         */
        public string GetComputerName()
        {
            return userCredential.computerName;
        }

        /*
         * Connection method
         */

        public void Connect(string nameSpace)
        {
            scopePath = $"\\\\{userCredential.computerName}\\{nameSpace}";
            scope = (isLocal) ? new ManagementScope(scopePath) : new ManagementScope(scopePath, credential);
            scope.Connect();
        }

        /*
         * Get WMI objects
         */
        public ManagementObjectCollection GetObjects(string className, string fields)
        {
            ObjectQuery query = new ObjectQuery($"SELECT {fields} FROM {className}");
            searcher = new ManagementObjectSearcher(scope, query);
            return searcher?.Get();
        }

        /*
         * Call WMI Method
         */
        public string MountPnPDeviceToPcip(string deviceID)
        {
            string DismountedDeviceInstancePath = String.Empty;

            ObjectQuery PnPQuery = new ObjectQuery($"SELECT DeviceId FROM Win32_PnPEntity WHERE DeviceID='{deviceID.Replace("\\", "\\\\")}'");
            ObjectQuery AssignDevService = new ObjectQuery("SELECT * FROM Msvm_AssignableDeviceService WHERE CreationClassName='Msvm_AssignableDeviceService'");

            Connect("root\\cimv2");
            using (var searcher = new ManagementObjectSearcher(scope, PnPQuery))
            {
                foreach (ManagementObject manageObject in searcher.Get().Cast<ManagementObject>())
                {
                    try
                    {
                        _ = (UInt32)manageObject.InvokeMethod("Disable", new object[] { });
                        break;
                    }
                    catch (IndexOutOfRangeException) { }
                }
            }

            Connect("root\\virtualization\\v2");
            using (var searcher = new ManagementObjectSearcher(scope, AssignDevService))
            {
                foreach (ManagementObject manageObject in searcher.Get().Cast<ManagementObject>())
                {
                    try
                    {
                        var dismountAssDev = manageObject.InvokeMethod("DismountAssignableDevice", new object[] { deviceID.Replace("\\", "\\\\") });

                        if (dismountAssDev is ManagementBaseObject mbo && mbo["DismountedDeviceInstancePath"] != null)
                        {
                            DismountedDeviceInstancePath = mbo["DismountedDeviceInstancePath"].ToString();
                        }

                        break;
                    }
                    catch (IndexOutOfRangeException) { }
                }
            }

            return DismountedDeviceInstancePath;
        }
    }
}
