using DDAGUI.Properties;
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
        public void ChangeGuestCacheType(string vmName, bool isEnableGuestControlCache)
        {
            Connect("root\\virtualization\\v2");

            ManagementObject vm = null;

            foreach (ManagementObject obj in GetObjects("Msvm_VirtualSystemSettingData", "*").Cast<ManagementObject>())
            {
                if (obj["Caption"].ToString().Equals("Virtual Machine Settings") && !obj["InstanceID"].ToString().Contains("Microsoft:Definition"))
                {
                    string hostName = obj["ElementName"]?.ToString();

                    if (hostName == null || hostName.Length == 0) continue;

                    if (hostName.Equals(vmName))
                    {
                        vm = obj;
                    }
                }
            }

            if (vm == null) throw new ManagementException("ChangeGuestCacheType: Unable to get the VM setting");

            vm["GuestControlledCacheTypes"] = (bool)isEnableGuestControlCache;

            ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_VirtualSystemManagementService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault() ?? throw new ManagementException("MountIntoVM: Assignment service is either not running or crashed");
            UInt32 outParams = (UInt32)srv.InvokeMethod("ModifySystemSettings", new object[]
            {
                vm.GetText(TextFormat.WmiDtd20)
            });

            switch (outParams)
            {
                case (UInt32)0:
                    break;
                case (UInt32)1:
                    throw new ManagementException("ChangeGuestCacheType: Not Supported");
                case (UInt32)2:
                    throw new ManagementException("ChangeGuestCacheType: Failed");
                case (UInt32)3:
                    throw new ManagementException("ChangeGuestCacheType: Timed out");
                case (UInt32)4:
                    throw new ManagementException("ChangeGuestCacheType: Invalid Parameter");
                case (UInt32)5:
                    throw new ManagementException("ChangeGuestCacheType: Invalid State");
                case (UInt32)6:
                    throw new ManagementException("ChangeGuestCacheType: Incompatible Parameters");
                case (UInt32)4096:
                    throw new ManagementException("ChangeGuestCacheType: Method Parameters Checked but failed to Execute");
                default:
                    throw new ManagementException("ChangeGuestCacheType: Unknown error");
            }
        }

        public void ChangeMemAllocate(string vmName, int highMem, int lowMem)
        {
            Connect("root\\virtualization\\v2");

            ManagementObject vm = null;

            foreach (ManagementObject obj in GetObjects("Msvm_VirtualSystemSettingData", "*").Cast<ManagementObject>())
            {
                if (obj["Caption"].ToString().Equals("Virtual Machine Settings") && !obj["InstanceID"].ToString().Contains("Microsoft:Definition"))
                {
                    string hostName = obj["ElementName"]?.ToString();

                    if (hostName == null || hostName.Length == 0) continue;

                    if (hostName.Equals(vmName))
                    {
                        vm = obj;
                    }
                }
            }

            if (vm == null) throw new ManagementException("ChangeMemAllocate: Unable to get the VM setting");

            vm["LowMmioGapSize"] = (UInt64)(lowMem);
            vm["HighMmioGapSize"] = (UInt64)(highMem);

            ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_VirtualSystemManagementService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault() ?? throw new ManagementException("MountIntoVM: Assignment service is either not running or crashed");
            UInt32 outParams = (UInt32)srv.InvokeMethod("ModifySystemSettings", new object[]
            {
                vm.GetText(TextFormat.WmiDtd20)
            });

            switch (outParams)
            {
                case (UInt32)0:
                    break;
                case (UInt32)1:
                    throw new ManagementException("ChangeMemAllocate: Not Supported");
                case (UInt32)2:
                    throw new ManagementException("ChangeMemAllocate: Failed");
                case (UInt32)3:
                    throw new ManagementException("ChangeMemAllocate: Timed out");
                case (UInt32)4:
                    throw new ManagementException("ChangeMemAllocate: Invalid Parameter");
                case (UInt32)5:
                    throw new ManagementException("ChangeMemAllocate: Invalid State");
                case (UInt32)6:
                    throw new ManagementException("ChangeMemAllocate: Incompatible Parameters");
                case (UInt32)4096:
                    throw new ManagementException("ChangeMemAllocate: Method Parameters Checked but failed to Execute");
                default:
                    throw new ManagementException("ChangeMemAllocate: Unknown error");
            }
        }

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
                    throw new ManagementException("MountPnPDeviceToPcip: Method Parameters Checked but failed to Execute");
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

        public void DismountPnPDeviceFromPcip(string devicePath)
        {
            Connect("root\\virtualization\\v2");
            string deviceLocation = string.Empty;
            string actualDevicePath = devicePath.Replace("PCIP\\", "PCI\\");

            foreach (ManagementObject obj in GetObjects("Msvm_PciExpress", "*").Cast<ManagementObject>())
            {
                if (obj["DeviceInstancePath"].ToString().Equals(devicePath))
                {
                    deviceLocation = obj["LocationPath"].ToString();
                    break;
                }
            }

            if (deviceLocation == string.Empty) throw new ManagementException("DismountPnPDeviceFromPcip: Unable to get the device location");

            ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_AssignableDeviceService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault() ?? throw new ManagementException("DismountPnPDeviceFromPcip: Assignment service is either not running or crashed");
            UInt32 virtv2_outparams = (UInt32)srv.InvokeMethod("MountAssignableDevice", new object[]
            {
                devicePath,
                deviceLocation
            });

            switch (virtv2_outparams)
            {
                case (UInt32)0:
                    break;
                case (UInt32)4096:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Method Parameters Checked but failed to Execute");
                case (UInt32)32768:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Failed");
                case (UInt32)32769:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Access Denied");
                case (UInt32)32770:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Not Supported");
                case (UInt32)32771:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Status is unknown");
                case (UInt32)32772:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Timeout");
                case (UInt32)32773:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Invalid parameter");
                case (UInt32)32774:
                    throw new ManagementException("DismountPnPDeviceFromPcip: System is in use");
                case (UInt32)32775:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Invalid state for this operation");
                case (UInt32)32776:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Incorrect data type");
                case (UInt32)32777:
                    throw new ManagementException("DismountPnPDeviceFromPcip: System is not available");
                case (UInt32)32778:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Out of memory");
                case (UInt32)32779:
                    throw new ManagementException("DismountPnPDeviceFromPcip: File not found");
                default:
                    throw new ManagementException("Unknow error in method DismountPnPDeviceFromPcip");
            }

            Connect("root\\cimv2");
            foreach (ManagementObject deviceSearcher in GetObjects("Win32_PnPEntity", "DeviceId").Cast<ManagementObject>())
            {
                if (deviceSearcher["DeviceId"].ToString().Equals(actualDevicePath))
                {
                    try
                    {
                        _ = (UInt32)deviceSearcher.InvokeMethod("Enable", new object[] { });
                    }
                    catch (IndexOutOfRangeException) { }
                    break;
                }
            }
        }

        public void MountIntoVM(string vmName, string deviceId)
        {
            Connect("root\\virtualization\\v2");

            ManagementObject vmCurrentSetting = null;
            ManagementObject setting = (new ManagementClass(scope, new ManagementPath("Msvm_PciExpressSettingData"), null)).CreateInstance();
            string hostRes = string.Empty, vmRes = string.Empty;

            foreach (ManagementObject device in GetObjects("Msvm_PciExpress", "*").Cast<ManagementObject>())
            {
                if (device["DeviceInstancePath"].ToString().Contains(deviceId.Replace("PCI\\", "PCIP\\")))
                {
                    string hostName = device["SystemName"].ToString();
                    string deviceIdPnP = device["DeviceId"].ToString().Replace("\\", "\\\\");
                    hostRes = $"\\\\{hostName}\\root\\virtualization\\v2:Msvm_PciExpress.CreationClassName=\"Msvm_PciExpress\",DeviceID=\"{deviceIdPnP}\",SystemCreationClassName=\"Msvm_ComputerSystem\",SystemName=\"{hostName}\"";
                    break;
                }
            }
            if (hostRes == String.Empty) throw new ManagementException("MountIntoVM: Unable to get HostResource");

            foreach (ManagementObject vmSetting in GetObjects("Msvm_VirtualSystemSettingData", "*").Cast<ManagementObject>())
            {
                if (vmSetting["Caption"].ToString().Equals("Virtual Machine Settings") && !vmSetting["InstanceID"].ToString().Contains("Microsoft:Definition"))
                {
                    string hostName = vmSetting["ElementName"]?.ToString();

                    if (hostName == null || hostName.Length == 0) continue;

                    if (hostName.Equals(vmName))
                    {
                        vmRes = $"{vmSetting["InstanceID"]}\\{Guid.NewGuid().ToString().ToUpper()}";
                        vmCurrentSetting = vmSetting;
                        break;
                    }

                }
            }
            if (vmRes == String.Empty) throw new ManagementException("MountIntoVM: Unable to generate InstanceID");
            if (vmCurrentSetting == null) throw new ManagementException("MountIntoVM: Unable to get setting for VM");

            setting["Address"] = String.Empty;
            setting["AddressOnParent"] = String.Empty;
            setting["AllocationUnits"] = "count";
            setting["AllowDirectTranslatedP2P"] = new bool[] { false };
            setting["AutomaticAllocation"] = true;
            setting["AutomaticDeallocation"] = true;
            setting["Caption"] = "PCI Express Port";
            setting["Connection"] = null;
            setting["ConsumerVisibility"] = (UInt16)3;
            setting["Description"] = "Microsoft Virtual PCI Express Port Setting Data";
            setting["ElementName"] = "PCI Express Port";
            setting["HostResource"] = new string[] { hostRes };
            setting["InstanceID"] = vmRes;
            setting["Limit"] = (UInt64)1;
            setting["MappingBehavior"] = null;
            setting["OtherResourceType"] = null;
            setting["Parent"] = null;
            setting["PoolID"] = String.Empty;
            setting["Reservation"] = (UInt64)1;
            setting["ResourceSubType"] = "Microsoft:Hyper-V:Virtual Pci Express Port";
            setting["ResourceType"] = (UInt16)32769;
            setting["TargetVtl"] = (uint)0;
            setting["VirtualFunctions"] = new UInt16[] { 0 };
            setting["VirtualQuantity"] = (UInt64)1;
            setting["VirtualSystemIdentifiers"] = new string[] { "{" + Guid.NewGuid() + "}" };
            setting["VirtualQuantityUnits"] = "count";
            setting["Weight"] = (UInt32)0;

            ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_VirtualSystemManagementService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault() ?? throw new ManagementException("MountIntoVM: Assignment service is either not running or crashed");

            UInt32 outParams = (UInt32)srv.InvokeMethod("AddResourceSettings", new object[]
            {
                vmCurrentSetting,
                new string[] { setting.GetText(TextFormat.WmiDtd20) }
            });

            switch (outParams)
            {
                case (UInt32)0:
                    break;
                case (UInt32)1:
                    throw new ManagementException("MountIntoVM: Not Supported");
                case (UInt32)2:
                    throw new ManagementException("MountIntoVM: Failed");
                case (UInt32)3:
                    throw new ManagementException("MountIntoVM: Timed out");
                case (UInt32)4:
                    throw new ManagementException("MountIntoVM: Invalid Parameter");
                case (UInt32)4096:
                    throw new ManagementException("MountIntoVM: Method Parameters Checked but failed to Execute");
                case (UInt32)4097:
                    throw new ManagementException("MountIntoVM: The function may not be called or is reserved for vendor");
                default:
                    throw new ManagementException("MountIntoVM: Unknown error");
            }
        }

        public void DismountFromVM(string deviceId)
        {
            Connect("root\\virtualization\\v2");

            ManagementObject[] deviceObj = new ManagementObject[1];

            foreach (ManagementObject obj in GetObjects("Msvm_PciExpressSettingData", "*").Cast<ManagementObject>())
            {
                if (obj["InstanceID"].ToString().Equals(deviceId))
                {
                    deviceObj[0] = obj;
                    break;
                }
            }

            if (deviceObj == null) throw new ManagementException("DismountFromVM: Error while finding the device ID");

            ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_VirtualSystemManagementService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault() ?? throw new ManagementException("DismountFromVM: Assignment service is either not running or crashed");
            UInt32 outParams = (UInt32)srv.InvokeMethod("RemoveResourceSettings", new object[]
            {
                deviceObj
            });

            switch (outParams)
            {
                case (UInt32)0:
                    break;
                case (UInt32)1:
                    throw new ManagementException("DismountFromVM: Not Supported");
                case (UInt32)2:
                    throw new ManagementException("DismountFromVM: Failed");
                case (UInt32)3:
                    throw new ManagementException("DismountFromVM: Timeout");
                case (UInt32)4:
                    throw new ManagementException("DismountFromVM: Invalid Parameter");
                case (UInt32)5:
                    throw new ManagementException("DismountFromVM: Invalid State");
                case (UInt32)6:
                    throw new ManagementException("DismountFromVM: Method Parameters Checked but failed to Execute");
                default:
                    throw new ManagementException("DismountFromVM: Unknown error");
            }
        }
    }
}