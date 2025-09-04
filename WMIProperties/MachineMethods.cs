using System;
using System.Linq;
using System.Management;
using System.Security;

namespace TheFlightSims.HyperVDPD.WMIProperties
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
            return (userCredential.computerName);
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
            return (searcher?.Get());
        }

        /*
         * Call WMI Method
         */
        public void ChangeGuestCacheType(string vmName, bool isEnableGuestControlCache)
        {
            Connect("root\\virtualization\\v2");

            UInt32 outParams = 32768;
            ManagementObject vm = null;

            foreach (ManagementObject obj in GetObjects("Msvm_VirtualSystemSettingData", "*").Cast<ManagementObject>())
            {
                string objCaption = obj["Caption"]?.ToString();
                string objInstanceId = obj["InstanceID"]?.ToString();

                if (objCaption == null || objInstanceId == null)
                {
                    obj.Dispose();
                    continue;
                }

                if (objCaption.Equals("Virtual Machine Settings") && !objInstanceId.Contains("Microsoft:Definition"))
                {
                    string hostName = obj["ElementName"]?.ToString();

                    if (hostName == null)
                    {
                        obj.Dispose();
                        continue;
                    }

                    if (hostName.Equals(vmName))
                    {
                        vm = obj;
                        break;
                    }
                }

                obj.Dispose();
            }

            if (vm == null)
            {
                throw new ManagementException("ChangeGuestCacheType: Unable to get the virtual machine setting");
            }

            try
            {
                vm["GuestControlledCacheTypes"] = isEnableGuestControlCache;

                ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_VirtualSystemManagementService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault() ?? throw new ManagementException("ChangeGuestCacheType: Assignment service is either not running or crashed");
                outParams = (UInt32)srv.InvokeMethod("ModifySystemSettings", new object[]
                {
                    vm.GetText(TextFormat.WmiDtd20)
                });
            }
            finally
            {
                vm?.Dispose();
            }

            switch (outParams)
            {
                case 0:
                    break;
                case 1:
                    throw new ManagementException("ChangeGuestCacheType: Not Supported");
                case 2:
                    throw new ManagementException("ChangeGuestCacheType: Failed");
                case 3:
                    throw new ManagementException("ChangeGuestCacheType: Timed out");
                case 4:
                    throw new ManagementException("ChangeGuestCacheType: Invalid Parameter");
                case 5:
                    throw new ManagementException("ChangeGuestCacheType: Invalid State");
                case 6:
                    throw new ManagementException("ChangeGuestCacheType: Incompatible Parameters");
                case 4096:
                    throw new ManagementException("ChangeGuestCacheType: Method Parameters Checked but failed to Execute");
                default:
                    throw new ManagementException($"ChangeGuestCacheType: Unknown error ({outParams})");
            }
        }

        public void ChangeMemAllocate(string vmName, UInt64 lowMem, UInt64 highMem)
        {
            Connect("root\\virtualization\\v2");

            ManagementObject vm = null;
            UInt32 outParams = 32768;

            foreach (ManagementObject obj in GetObjects("Msvm_VirtualSystemSettingData", "*").Cast<ManagementObject>())
            {
                string objCaption = obj["Caption"]?.ToString();
                string objInstanceId = obj["InstanceID"]?.ToString();

                if (objCaption == null || objInstanceId == null)
                {
                    obj.Dispose();
                    continue;
                }

                if (objCaption.Equals("Virtual Machine Settings") && !objInstanceId.Contains("Microsoft:Definition"))
                {
                    string hostName = obj["ElementName"]?.ToString();

                    if (hostName == null)
                    {
                        obj.Dispose();
                        continue;
                    }

                    if (hostName.Equals(vmName))
                    {
                        vm = obj;
                    }
                }

                obj.Dispose();
            }

            if (vm == null)
            {
                throw new ManagementException("ChangeMemAllocate: Unable to get the virtual machine setting");
            }

            try
            {
                vm["LowMmioGapSize"] = lowMem;
                vm["HighMmioGapSize"] = highMem;

                ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_VirtualSystemManagementService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault() ?? throw new ManagementException("ChangeMemAllocate: Assignment service is either not running or crashed");
                outParams = (UInt32)srv.InvokeMethod("ModifySystemSettings", new object[]
                {
                    vm.GetText(TextFormat.WmiDtd20)
                });
            }
            finally
            {
                vm?.Dispose();
            }

            switch (outParams)
            {
                case 0:
                    break;
                case 1:
                    throw new ManagementException("ChangeMemAllocate: Not Supported");
                case 2:
                    throw new ManagementException("ChangeMemAllocate: Failed");
                case 3:
                    throw new ManagementException("ChangeMemAllocate: Timed out");
                case 4:
                    throw new ManagementException("ChangeMemAllocate: Invalid Parameter");
                case 5:
                    throw new ManagementException("ChangeMemAllocate: Invalid State");
                case 6:
                    throw new ManagementException("ChangeMemAllocate: Incompatible Parameters");
                case 4096:
                    throw new ManagementException("ChangeMemAllocate: Method Parameters Checked but failed to Execute");
                default:
                    throw new ManagementException($"ChangeMemAllocate: Unknown error ({outParams})");
            }
        }

        public void ChangePnpDeviceBehaviour(string deviceID, string behaviourMethod)
        {
            if (behaviourMethod == "Disable" || behaviourMethod == "Enable")
            {
                Connect("root\\cimv2");
                foreach (ManagementObject deviceSearcher in GetObjects("Win32_PnPEntity", "DeviceId").Cast<ManagementObject>())
                {
                    string devSearcherId = deviceSearcher["DeviceId"]?.ToString();

                    if (devSearcherId == null)
                    {
                        deviceSearcher.Dispose();
                        continue;
                    }

                    if (devSearcherId.Equals(deviceID))
                    {
                        try
                        {
                            _ = (UInt32)deviceSearcher.InvokeMethod(behaviourMethod, new object[] { });
                        }
                        catch (IndexOutOfRangeException) { }
                        finally
                        {
                            deviceSearcher.Dispose();
                        }

                        break;
                    }

                    deviceSearcher.Dispose();
                }
            }
        }

        public void MountPnPDeviceToPcip(string deviceID)
        {
            Connect("root\\virtualization\\v2");

            UInt32 outObj = 32768;

            ManagementObject setting = ((new ManagementClass(scope, new ManagementPath("Msvm_AssignableDeviceDismountSettingData"), null))?.CreateInstance()) ?? throw new ManagementException("MountPnPDeviceToPcip: Unable to get the setting class");
            ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_AssignableDeviceService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault() ?? throw new ManagementException("MountPnPDeviceToPcip: Assignment service is either not running or crashed");

            try
            {
                // Msvm_AssignableDeviceDismountSettingData
                setting["DeviceInstancePath"] = (string.IsNullOrEmpty(deviceID)) ? string.Empty : deviceID;
                setting["DeviceLocationPath"] = string.Empty;
                setting["RequireAcsSupport"] = false;
                setting["RequireDeviceMitigations"] = false;

                outObj = (UInt32)srv.InvokeMethod("DismountAssignableDevice", new object[] {
                    setting.GetText(TextFormat.WmiDtd20)
                });
            }
            finally
            {
                srv?.Dispose();
                setting?.Dispose();
            }

            switch (outObj)
            {
                case 0:
                    break;
                case 4096:
                    throw new ManagementException("MountPnPDeviceToPcip: Method Parameters Checked but failed to Execute");
                case 32768:
                    throw new ManagementException("MountPnPDeviceToPcip: Access Denied");
                case 32770:
                    throw new ManagementException("MountPnPDeviceToPcip: Not Supported");
                case 32771:
                    throw new ManagementException("MountPnPDeviceToPcip: Status is unknown");
                case 32772:
                    throw new ManagementException("MountPnPDeviceToPcip: Timeout");
                case 32773:
                    throw new ManagementException("MountPnPDeviceToPcip: Invalid parameter");
                case 32774:
                    throw new ManagementException("MountPnPDeviceToPcip: System is in use");
                case 32775:
                    throw new ManagementException("MountPnPDeviceToPcip: Invalid state for this operation");
                case 32776:
                    throw new ManagementException("MountPnPDeviceToPcip: Incorrect data type");
                case 32777:
                    throw new ManagementException("MountPnPDeviceToPcip: System is not available");
                case 32778:
                    throw new ManagementException("MountPnPDeviceToPcip: Out of memory");
                case 32779:
                    throw new ManagementException("MountPnPDeviceToPcip: File not found");
                default:
                    throw new ManagementException($"Unknow error in method MountPnPDeviceToPcip ({outObj})");
            }
        }

        public void DismountPnPDeviceFromPcip(string devicePath)
        {
            Connect("root\\virtualization\\v2");

            string deviceLocation = string.Empty;
            UInt32 virtv2_outparams = 32768;

            foreach (ManagementObject obj in GetObjects("Msvm_PciExpress", "*").Cast<ManagementObject>())
            {
                string devInstancePath = obj["DeviceInstancePath"]?.ToString();

                if (devInstancePath == null)
                {
                    obj.Dispose();
                    continue;
                }

                if (devInstancePath.Equals(devicePath))
                {
                    string devLoc = obj["LocationPath"]?.ToString();

                    if (devLoc != null)
                    {
                        deviceLocation = devLoc;
                        obj.Dispose();
                        break;
                    }
                }

                obj.Dispose();
            }

            if (deviceLocation == string.Empty)
            {
                throw new ManagementException("DismountPnPDeviceFromPcip: Unable to get the device location");
            }

            ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_AssignableDeviceService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault() ?? throw new ManagementException("DismountPnPDeviceFromPcip: Assignment service is either not running or crashed");

            try
            {
                virtv2_outparams = (UInt32)srv.InvokeMethod("MountAssignableDevice", new object[]
                {
                    devicePath,
                    deviceLocation
                });
            }
            finally
            {
                srv?.Dispose();
            }

            switch (virtv2_outparams)
            {
                case 0:
                    break;
                case 4096:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Method Parameters Checked but failed to Execute");
                case 32768:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Failed");
                case 32769:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Access Denied");
                case 32770:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Not Supported");
                case 32771:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Status is unknown");
                case 32772:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Timeout");
                case 32773:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Invalid parameter");
                case 32774:
                    throw new ManagementException("DismountPnPDeviceFromPcip: System is in use");
                case 32775:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Invalid state for this operation");
                case 32776:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Incorrect data type");
                case 32777:
                    throw new ManagementException("DismountPnPDeviceFromPcip: System is not available");
                case 32778:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Out of memory");
                case 32779:
                    throw new ManagementException("DismountPnPDeviceFromPcip: File not found");
                default:
                    throw new ManagementException($"Unknow error in method DismountPnPDeviceFromPcip ({virtv2_outparams})");
            }
        }

        public void MountIntoVM(string vmName, string deviceId)
        {
            Connect("root\\virtualization\\v2");

            ManagementObject vmCurrentSetting = null;
            ManagementObject setting = (new ManagementClass(scope, new ManagementPath("Msvm_PciExpressSettingData"), null)).CreateInstance();
            ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_VirtualSystemManagementService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault() ?? throw new ManagementException("MountIntoVM: Assignment service is either not running or crashed");

            UInt32 outParams = 32768;
            string hostRes = string.Empty, vmRes = string.Empty;

            foreach (ManagementObject device in GetObjects("Msvm_PciExpress", "*").Cast<ManagementObject>())
            {
                string devInstancePath = device["DeviceInstancePath"]?.ToString();

                if (devInstancePath == null)
                {
                    device.Dispose();
                    continue;
                }

                if (devInstancePath.Contains(deviceId.Replace("PCI\\", "PCIP\\")))
                {
                    string hostName = device["SystemName"]?.ToString();
                    string deviceIdPnP = device["DeviceId"]?.ToString();

                    if (hostName != null || deviceIdPnP != null)
                    {
                        deviceIdPnP = deviceIdPnP.Replace("\\", "\\\\");
                        hostRes = $"\\\\{hostName}\\root\\virtualization\\v2:Msvm_PciExpress.CreationClassName=\"Msvm_PciExpress\",DeviceID=\"{deviceIdPnP}\",SystemCreationClassName=\"Msvm_ComputerSystem\",SystemName=\"{hostName}\"";
                        device.Dispose();
                        break;
                    }

                    device.Dispose();
                }
            }
            if (hostRes == String.Empty)
            {
                throw new ManagementException("MountIntoVM: Unable to get HostResource");
            }

            foreach (ManagementObject vmSetting in GetObjects("Msvm_VirtualSystemSettingData", "*").Cast<ManagementObject>())
            {
                string vmCaption = vmSetting["Caption"]?.ToString();
                string vmInstanceID = vmSetting["InstanceID"]?.ToString();

                if (vmCaption == null || vmInstanceID == null)
                {
                    vmSetting.Dispose();
                    continue;
                }

                if (vmCaption.Equals("Virtual Machine Settings") && !vmInstanceID.Contains("Microsoft:Definition"))
                {
                    string hostName = vmSetting["ElementName"]?.ToString();

                    if (hostName == null)
                    {
                        vmSetting.Dispose();
                        continue;
                    }

                    if (hostName.Equals(vmName))
                    {
                        vmRes = $"{vmSetting["InstanceID"]}\\{Guid.NewGuid().ToString().ToUpper()}";
                        vmCurrentSetting = vmSetting;
                        vmSetting.Dispose();
                        break;
                    }

                }
            }
            if (vmRes == String.Empty)
            {
                throw new ManagementException("MountIntoVM: Unable to generate InstanceID");
            }
            if (vmCurrentSetting == null)
            {
                throw new ManagementException("MountIntoVM: Unable to get setting for VM");
            }

            try
            {
                // CIM_SettingData
                setting["Caption"] = "PCI Express Port";
                setting["Description"] = "Microsoft Virtual PCI Express Port Setting Data";
                setting["InstanceID"] = vmRes;
                setting["ElementName"] = "PCI Express Port";

                // CIM_ResourceAllocationSettingData
                setting["ResourceType"] = (UInt16)32769;
                setting["OtherResourceType"] = null;
                setting["ResourceSubType"] = "Microsoft:Hyper-V:Virtual Pci Express Port";
                setting["PoolID"] = String.Empty;
                setting["ConsumerVisibility"] = (UInt16)3;
                setting["HostResource"] = new string[] { hostRes };
                setting["AllocationUnits"] = "count";
                setting["VirtualQuantity"] = (UInt64)1;
                setting["Reservation"] = (UInt64)1;
                setting["Limit"] = (UInt64)1;
                setting["Weight"] = (UInt32)0;
                setting["AutomaticAllocation"] = true;
                setting["AutomaticDeallocation"] = true;
                setting["Parent"] = null;
                setting["Connection"] = null;
                setting["Address"] = String.Empty;
                setting["MappingBehavior"] = null;
                setting["AddressOnParent"] = String.Empty;
                setting["VirtualQuantityUnits"] = "count";

                // Msvm_PciExpressSettingData
                setting["VirtualFunctions"] = new UInt16[] { 0 };
                setting["VirtualSystemIdentifiers"] = new string[] { "{" + Guid.NewGuid() + "}" };


                outParams = (UInt32)srv.InvokeMethod("AddResourceSettings", new object[]
                {
                    vmCurrentSetting,
                    new string[] { setting.GetText(TextFormat.WmiDtd20) }
                });
            }
            finally
            {
                setting?.Dispose();
                srv?.Dispose();
            }

            switch (outParams)
            {
                case 0:
                    break;
                case 1:
                    throw new ManagementException("MountIntoVM: Not Supported");
                case 2:
                    throw new ManagementException("MountIntoVM: Failed");
                case 3:
                    throw new ManagementException("MountIntoVM: Timed out");
                case 4:
                    throw new ManagementException("MountIntoVM: Invalid Parameter");
                case 4096:
                    throw new ManagementException("MountIntoVM: Method Parameters Checked but failed to Execute");
                case 4097:
                    throw new ManagementException("MountIntoVM: The function may not be called or is reserved for vendor");
                default:
                    throw new ManagementException($"MountIntoVM: Unknown error ({outParams})");
            }
        }

        public void DismountFromVM(string deviceId)
        {
            Connect("root\\virtualization\\v2");

            ManagementObject[] deviceObj = new ManagementObject[1];
            ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_VirtualSystemManagementService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault() ?? throw new ManagementException("DismountFromVM: Assignment service is either not running or crashed");
            UInt32 outParams = 32768;

            foreach (ManagementObject obj in GetObjects("Msvm_PciExpressSettingData", "*").Cast<ManagementObject>())
            {
                string objInstanceId = obj["InstanceID"]?.ToString();

                if (objInstanceId != null)
                {
                    if (objInstanceId.Equals(deviceId))
                    {
                        deviceObj[0] = obj;
                        obj.Dispose();
                        break;
                    }
                }

                obj.Dispose();
            }

            if (deviceObj[0] == null)
            {
                throw new ManagementException("DismountFromVM: Error while finding the device ID");
            }

            try
            {
                outParams = (UInt32)srv.InvokeMethod("RemoveResourceSettings", new object[]
                {
                    deviceObj
                });
            }
            finally
            {
                srv?.Dispose();
            }

            switch (outParams)
            {
                case 0:
                    break;
                case 1:
                    throw new ManagementException("DismountFromVM: Not Supported");
                case 2:
                    throw new ManagementException("DismountFromVM: Failed");
                case 3:
                    throw new ManagementException("DismountFromVM: Timeout");
                case 4:
                    throw new ManagementException("DismountFromVM: Invalid Parameter");
                case 5:
                    throw new ManagementException("DismountFromVM: Invalid State");
                case 6:
                    throw new ManagementException("DismountFromVM: Method Parameters Checked but failed to Execute");
                default:
                    throw new ManagementException($"DismountFromVM: Unknown error ({outParams})");
            }
        }
    }
}