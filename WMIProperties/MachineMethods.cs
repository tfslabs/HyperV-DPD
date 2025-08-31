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
        public void MountPnPDeviceToPcip(string deviceID)
        {
            UInt32 outObj = (UInt32)32779;
            Connect("root\\cimv2");
            foreach (ManagementObject deviceSearcher in GetObjects("Win32_PnPEntity", "DeviceId").Cast<ManagementObject>())
            {
                if (deviceSearcher["DeviceId"].ToString().Equals(deviceID))
                {
                    try
                    {
                        _ = (UInt32)deviceSearcher.InvokeMethod("Disable", new object[] { });
                    }
                    catch (IndexOutOfRangeException) { }
                    break;
                }
            }

            Connect("root\\virtualization\\v2");
            ManagementObject setting = (new ManagementClass(scope, new ManagementPath("Msvm_AssignableDeviceDismountSettingData"), null)).CreateInstance();

            setting["DeviceInstancePath"] = deviceID;
            setting["DeviceLocationPath"] = string.Empty;
            setting["RequireAcsSupport"] = false;
            setting["RequireDeviceMitigations"] = false;

            foreach (ManagementObject svc in GetObjects("Msvm_AssignableDeviceService", "*").Cast<ManagementObject>())
            {
                outObj = (UInt32)svc.InvokeMethod("DismountAssignableDevice", new object[] { setting.GetText(TextFormat.WmiDtd20) });
                break;
            }

            switch (outObj)
            {
                case (UInt32)0:
                    break;
                case (UInt32)4096:
                    throw new ManagementException("MountPnPDeviceToPcip: Method Parameters Checked - Job Started");
                case (UInt32)32768:
                    throw new ManagementException("MountPnPDeviceToPcip: Access Denied");
                case (UInt32)32770:
                    throw new ManagementException("MountPnPDeviceToPcip: Not Supported");
                case (UInt32)32771:
                    throw new ManagementException("MountPnPDeviceToPcip: Status is unknown");
                case (UInt32)32772:
                    throw new ManagementException("MountPnPDeviceToPcip: Timeout");
                case (UInt32)32773:
                    throw new ManagementException("MountPnPDeviceToPcip: Invalid parameter");
                case (UInt32)32774:
                    throw new ManagementException("MountPnPDeviceToPcip: System is in use");
                case (UInt32)32775:
                    throw new ManagementException("MountPnPDeviceToPcip: Invalid state for this operation");
                case (UInt32)32776:
                    throw new ManagementException("MountPnPDeviceToPcip: Incorrect data type");
                case (UInt32)32777:
                    throw new ManagementException("MountPnPDeviceToPcip: System is not available");
                case (UInt32)32778:
                    throw new ManagementException("MountPnPDeviceToPcip: Out of memory");
                case (UInt32)32779:
                    throw new ManagementException("MountPnPDeviceToPcip: File not found");
                default:
                    throw new ManagementException("Unknow error in method MountPnPDeviceToPcip");
            }
        }
    }
}