using System;
using System.Linq;
using System.Management;
using System.Security;

/*
 * The default WMI method class includes:
 *  1. MachineMethods.cs - A base class for WMI operations on local and remote machines
 */
namespace TheFlightSims.HyperVDPD.WMIProperties
{
    /*
     * Machine Method class
     * It is used to connect and perform WMI tasks to local and remote machines through WMI
     * 
     * Critial Note: There are two states of the PCI device:
     *  1. PCI device in root\cimv2. These PCI devices are unmounted and still under control of host OS
     *  2. PCI device in root\virtualization\V2. These PCI devices are MOUNTED into PCIP and under the control of hypervisor
     * 
     * Security Note:
     *  1. When connecting to remote machines, ensure that the credentials are handled securely.
     *  2. Use SecureString for passwords to enhance security.
     *  3. The application should run with appropriate permissions to access WMI on the target machines.
     *      For example, locally it may require administrative privileges and do not run with external credentials.
     */
    public class MachineMethods
    {
        ////////////////////////////////////////////////////////////////
        /// Global Properties and Constructors Region
        ///     This region contains global properties and constructors 
        ///     for the MachineMethods class.
        ////////////////////////////////////////////////////////////////

        /*
         * Global properties
         *  userCredential: A tuple to store computer name, username, and password
         *  credential: A ConnectionOptions object to store WMI connection options
         *  scope: A ManagementScope object to define the WMI scope
         *  searcher: A ManagementObjectSearcher object to perform WMI queries
         *  scopePath: A string to store the WMI scope path
         *  isLocal: A boolean to indicate whether the connection is local or remote
         */

        protected (string computerName, string username, SecureString password) userCredential;
        protected ConnectionOptions credential;
        protected ManagementScope scope;
        protected ManagementObjectSearcher searcher;

        protected string scopePath;
        protected bool isLocal;

        /*
         * Constructor for local connection
         *  Note: Local connections do not require username and password. 
         *  It uses current user context.
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

        /*
         * Constructor for remote connection
         *  Note: Remote connections require username and password.
         */
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

        ////////////////////////////////////////////////////////////////
        ///  Connect to a name space region
        ///     This region contains methods to connect to a WMI name space.
        ///     For example, "root\\cimv2".
        ////////////////////////////////////////////////////////////////

        public void Connect(string nameSpace)
        {
            scopePath = $"\\\\{userCredential.computerName}\\{nameSpace}";
            scope = (isLocal) ? new ManagementScope(scopePath) : new ManagementScope(scopePath, credential);
            scope.Connect();
        }

        ////////////////////////////////////////////////////////////////
        /// Get methods region (connection setting, not the WMI data)
        ///     Various get methods to retrieve information about the machine
        ///     For example, computer name, current user name, etc.
        ////////////////////////////////////////////////////////////////

        /*
         * Get methods for computer name 
         */
        public string GetComputerName()
        {
            return userCredential.computerName;
        }

        /*
         * Get methods for current user name
         */
        public string GetUsername()
        {
            return userCredential.username;
        }

        ////////////////////////////////////////////////////////////////
        /// Get methods region (for WMI data)
        /// This region contains methods to retrieve WMI objects from the connected scope.
        ////////////////////////////////////////////////////////////////

        public ManagementObjectCollection GetObjects(string className, string fields)
        {
            ObjectQuery query = new ObjectQuery($"SELECT {fields} FROM {className}");
            searcher = new ManagementObjectSearcher(scope, query);
            return (searcher?.Get());
        }

        ////////////////////////////////////////////////////////////////
        /// Set methods region (for WMI object)
        /// This region contains methods to modify WMI objects in the connected scope.
        ////////////////////////////////////////////////////////////////

        /*
         * Change Guest Cache Type Method
         * This method modifies the GuestControlledCacheTypes property of a virtual machine's settings.
         * 
         * How does it work?
         *   1. It connects to the Hyper-V WMI namespace (root\virtualization\v2)
         *   2. From variable vmName, get the VM Management Object. Note: Check if the VM exists 
         *      after assignment
         *   3. From the virtualization service, it invokes the ModifySystemSettings method to change the 
         *      GuestControlledCacheTypes property. See in the references for more details.
         *   4. From the out parameter, check the return value and compare it into Success or Failure state
         *   
         *  Note:
         *      - For each value of the get ManagementObject, it is required to check for null value
         *      - It is required to get the out parameter to check the status of the operation.
         *      - No matter how the service operate with the parameter, the virtual machine instance must 
         *          be disposed
         *      
         *  References: 
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-virtualsystemsettingdata
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-virtualsystemmanagementservice
         *  
         *  Limitations:
         *      - The method assumes that the virtual machine with the specified name exists.
         *      - Windows 8/Windows Server 2012 or later is required for this method to work.
         *      - WMI Filtering may affect the operation.
         */
        public void ChangeGuestCacheType(string vmName, bool isEnableGuestControlCache)
        {
            // 1. Connect to the namespace
            Connect("root\\virtualization\\v2");

            // Set the property
            uint outParams = 32768;
            ManagementObject vm = null;

            // 2. Get the target ManagementObject from the vmName
            using (ManagementObjectCollection vmSettingDataList = GetObjects("Msvm_VirtualSystemSettingData", "*"))
            {
                foreach (ManagementObject obj in vmSettingDataList.Cast<ManagementObject>())
                {
                    /*
                     * In this loop, it first check if the target VM setting is the actual setting, and not the definition
                     * After that, if the ElementName of the ManagementObject matches with the vmName param, set the target
                     *  object to that ManagementObject
                     *  
                     *  Note: Always dispose on null values
                     */

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
            }

            if (vm == null)
            {
                throw new ManagementException("ChangeGuestCacheType: Unable to get the virtual machine setting");
            }

            // 3. Modify for the value
            try
            {
                vm["GuestControlledCacheTypes"] = isEnableGuestControlCache;

                // Get the service object
                ManagementObject srv = new ManagementClass(
                        scope,
                        new ManagementPath("Msvm_VirtualSystemManagementService"),
                        null
                    ).GetInstances().Cast<ManagementObject>().FirstOrDefault();

                // If the service is null, it is crashed
                if (srv == null)
                {
                    throw new ManagementException(
                        "ChangeGuestCacheType: Assignment service is either not running or crashed"
                    );
                }

                // Casting and get the return value of the Msvm_VirtualSystemManagementService
                outParams = (uint)srv.InvokeMethod("ModifySystemSettings", new object[]
                {
                    vm.GetText(TextFormat.WmiDtd20)
                });
            }
            finally
            {
                vm?.Dispose();
            }

            // 4. Check for the out param. Returns 0 in case success, other numbers as failure, and failback as unknown
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
                    throw new ManagementException("ChangeGuestCacheType: Method Parameters Checked - Job Started");
                default:
                    throw new ManagementException($"ChangeGuestCacheType: Unknown error ({outParams})");
            }
        }

        /*
         * Change Memory Allocation
         * This method changes the low memory allocation and high memory allocation
         * 
         * How does it work?
         *   1. It connects to the Hyper-V WMI namespace (root\virtualization\v2)
         *   2. From variable vmName, get the VM Management Object. Note: Check if the VM exists 
         *      after assignment
         *   3. From the virtualization service, it invokes the ModifySystemSettings method to change the 
         *      LowMmioGapSize and HighMmioGapSize properties. See in the references for more details.
         *   4. From the out parameter, check the return value and compare it into Success or Failure state
         *   
         *  Note:
         *      - For each value of the get ManagementObject, it is required to check for null value
         *      - It is required to get the out parameter to check the status of the operation.
         *      - No matter how the service operate with the parameter, the virtual machine instance must 
         *          be disposed
         *      
         *  References: 
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-virtualsystemsettingdata
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-virtualsystemmanagementservice
         *  
         *  Limitations:
         *      - The method assumes that the virtual machine with the specified name exists.
         *      - Windows 8/Windows Server 2012 or later is required for this method to work.
         *      - WMI Filtering may affect the operation.
         */
        public void ChangeMemAllocate(string vmName, ulong lowMem, ulong highMem)
        {
            // 1. Connect to the namespace
            Connect("root\\virtualization\\v2");

            // Set the property
            uint outParams = 32768;
            ManagementObject vm = null;

            // 2. Get the target ManagementObject from the vmName
            using (ManagementObjectCollection vmSettingDataList = GetObjects("Msvm_VirtualSystemSettingData", "*"))
            {
                foreach (ManagementObject obj in vmSettingDataList.Cast<ManagementObject>())
                {
                    /*
                     * In this loop, it first check if the target VM setting is the actual setting, and not the definition
                     * After that, if the ElementName of the ManagementObject matches with the vmName param, set the target
                     *  object to that ManagementObject
                     *  
                     *  Note: Always dispose on null values
                     */
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
            }
            
            if (vm == null)
            {
                throw new ManagementException("ChangeMemAllocate: Unable to get the virtual machine setting");
            }

            // 3. Modify for the value
            try
            {
                vm["LowMmioGapSize"] = lowMem;
                vm["HighMmioGapSize"] = highMem;

                // Get the management service
                ManagementObject srv = new ManagementClass(
                    scope,
                    new ManagementPath("Msvm_VirtualSystemManagementService"), null
                ).GetInstances().Cast<ManagementObject>().FirstOrDefault();

                // If null, the service is crashed
                if (srv == null)
                {
                    throw new ManagementException("ChangeMemAllocate: Assignment service is either not running or crashed");
                }

                // Casting and get the return value of the Msvm_VirtualSystemManagementService
                outParams = (uint)srv.InvokeMethod("ModifySystemSettings", new object[]
                {
                    vm.GetText(TextFormat.WmiDtd20)
                });
            }
            finally
            {
                vm?.Dispose();
            }

            // 4. Check for the out param. Returns 0 in case success, other numbers as failure, and failback as unknown
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
                    throw new ManagementException("ChangeMemAllocate: Method Parameters Checked - Job Started");
                default:
                    throw new ManagementException($"ChangeMemAllocate: Unknown error ({outParams})");
            }
        }

        /*
         * Change the PCI device behaviour
         * This method enable/disable the target PCI device on the host machine
         * 
         * How does it work?
         *   1. It connects to the CIMv2 namespace (root\cimv2)
         *   2. From the provided the deviceID (of PCI, not PCIP), try to invoke method "Disable" or
         *      "Enable", depends on the input parameter isDisable. Note: the device will report 
         *      IndexOutOfRangeException, just ignore it.
         *   
         *  Note:
         *      - For each value of the get ManagementObject, it is required to check for null value
         *      - No matter how the service operate with the parameter, any WMI object must be disposed
         *      
         *  References: 
         *      - https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/win32-pnpentity#methods
         *  
         *  Limitations:
         *      - Windows Vista/Windows Server 2008 or later is required for this method to work.
         *      - WMI Filtering may affect the operation.
         */
        public void ChangePnpDeviceBehaviour(string deviceID, bool isDisable)
        {
            // 1. Connect to the root name space
            Connect("root\\cimv2");

            // 2. From the device list, try to find the device that matches the deviceID
            using (ManagementObjectCollection pnpEntityList = GetObjects("Win32_PnPEntity", "DeviceId"))
            {
                foreach (ManagementObject deviceSearcher in pnpEntityList.Cast<ManagementObject>())
                {
                    // Get the device ID
                    string devSearcherId = deviceSearcher["DeviceId"]?.ToString();

                    // Dispose the device when find null
                    if (devSearcherId == null)
                    {
                        deviceSearcher.Dispose();
                        continue;
                    }

                    // Check if the target device is 
                    if (devSearcherId.Equals(deviceID))
                    {
                        try
                        {
                            _ = (uint)deviceSearcher.InvokeMethod(
                                (isDisable) ? "Disable" : "Enable",
                                new object[] { }
                            );
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

        /*
         * Mount PnP Device into PCIP
         * This method mounts a device from PCI (root\cimv2) into PCIP (root\virtualization\v2)
         * 
         * How does it work?
         *   1. It connects to the Hyper-V namespace (root\virtualization\v2)
         *   2. From the device ID, try to mount the device into the PCIP using the service Msvm_AssignableDeviceService
         *   3. From the out parameter, check the return value and compare it into Success or Failure state
         *   
         *  Note:
         *      - For each value of the get ManagementObject, it is required to check for null value
         *      - No matter how the service operate with the parameter, any WMI object must be disposed
         *      
         *  References: 
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-assignabledevicedismountsettingdata
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-assignabledeviceservice
         *  
         *  Limitations:
         *      - Windows 10 1703/Windows Server 2016 or later is required for this method to work.
         *      - WMI Filtering may affect the operation.
         */
        public void MountPnPDeviceToPcip(string deviceID)
        {
            // 1. Connect to the Hyper-V WMI namespace
            Connect("root\\virtualization\\v2");

            // Set for the out param
            uint outObj = 32768;

            // Create new object as the input param from the Msvm_AssignableDeviceDismountSettingData for the WMI method
            ManagementObject setting = new ManagementClass(
                scope,
                new ManagementPath("Msvm_AssignableDeviceDismountSettingData"),
                null
             )?.CreateInstance();

            if (setting == null)
            {
                throw new ManagementException("MountPnPDeviceToPcip: Unable to get the setting class");
            }

            // Get the Assignable Device Service WMI object
            ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_AssignableDeviceService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault();

            if (srv == null)
            {
                throw new ManagementException("MountPnPDeviceToPcip: Assignment service is either not running or crashed");
            }

            // 2. Mount the device into the PCIP.
            try
            {
                setting["DeviceInstancePath"] = deviceID;
                setting["DeviceLocationPath"] = string.Empty;
                setting["RequireAcsSupport"] = false;
                setting["RequireDeviceMitigations"] = false;

                // Call for service method
                outObj = (uint)srv.InvokeMethod(
                    "DismountAssignableDevice",
                    new object[] {
                    setting.GetText(TextFormat.WmiDtd20)
                });
            }
            finally
            {
                srv?.Dispose();
                setting?.Dispose();
            }

            // 3. Check for the out param. Returns 0 in case success, other numbers as failure, and failback as unknown
            switch (outObj)
            {
                case 0:
                    break;
                case 4096:
                    throw new ManagementException("MountPnPDeviceToPcip: Method Parameters Checked - Job Started");
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

        /*
         * Dismount PnP Device from PCIP
         * This method dismounts a device from PCIP (root\virtualization\v2)
         * 
         * How does it work?
         *   1. It connects to the Hyper-V namespace (root\virtualization\v2)
         *   2. From devicePath variable, get the location path
         *   3. From the devicePath and PCIP device location path, try to dismount the
         *      device using Msvm_AssignableDeviceService and get the out param
         *   4. From the out parameter, check the return value and compare it into Success or 
         *      Failure state
         *   
         *  Note:
         *      - For each value of the get ManagementObject, it is required to check for null value
         *      - No matter how the service operate with the parameter, any WMI object must be disposed
         *      
         *  References: 
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-assignabledeviceservice
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-pciexpress
         *  
         *  Limitations:
         *      - Windows 10 1703/Windows Server 2016 or later is required for this method to work.
         *      - WMI Filtering may affect the operation.
         */
        public void DismountPnPDeviceFromPcip(string devicePath)
        {
            // 1. Connect to Hyper-V WMI and 
            Connect("root\\virtualization\\v2");

            // Set the out param and device location
            uint outObj = 32768;
            string deviceLocation = string.Empty;

            // 2. Get the device location path
            using (ManagementObjectCollection pciExpressList = GetObjects("Msvm_PciExpress", "*"))
            {
                foreach (ManagementObject obj in pciExpressList.Cast<ManagementObject>())
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
            }
            
            // In case not found, throw exception
            if (deviceLocation == string.Empty)
            {
                throw new ManagementException("DismountPnPDeviceFromPcip: Unable to get the device location");
            }

            // 3. Call the service 
            ManagementObject srv = new ManagementClass(
                scope,
                new ManagementPath("Msvm_AssignableDeviceService"),
                null
            ).GetInstances().Cast<ManagementObject>().FirstOrDefault();

            if (srv == null)
            {
                throw new ManagementException(
                    "DismountPnPDeviceFromPcip: Assignment service is either not running or crashed"
                );
            }

            try
            {
                outObj = (uint)srv.InvokeMethod("MountAssignableDevice", new object[]
                {
                    devicePath,
                    deviceLocation
                });
            }
            finally
            {
                srv?.Dispose();
            }

            // 4. Check for the out param. Returns 0 in case success, other numbers as failure, and failback as unknown
            switch (outObj)
            {
                case 0:
                    break;
                case 4096:
                    throw new ManagementException("DismountPnPDeviceFromPcip: Method Parameters Checked - Job Started");
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
                    throw new ManagementException($"Unknow error in method DismountPnPDeviceFromPcip ({outObj})");
            }
        }

        /*
         * Mount PCIP device into the VM (Must be invoked when it is still in PCI)
         * 
         * How does it work?
         *   1. It connects to the Hyper-V namespace (root\virtualization\v2)
         *   2. Then, creates a new instance of PCI Express Setting Data from Msvm_PciExpressSettingData
         *   3. From the PCIP device list, find if there is any device matches the PCI that has mounted
         *   4. Get the VM setting to mount new device into the VM
         *   5. Now try to mount the VM using vmCurrentSetting and setting
         *   
         *  Note:
         *      - For each value of the get ManagementObject, it is required to check for null value
         *      - No matter how the service operate with the parameter, any WMI object must be disposed
         *      - The function must be invoked from the PCI, that means the when the device is still with
         *      
         *  References: 
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-pciexpresssettingdata
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-virtualsystemmanagementservice
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-pciexpress
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-virtualsystemsettingdata
         *      - https://learn.microsoft.com/en-us/previous-versions/cc136911(v=vs.85)
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/cim-resourceallocationsettingdata
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-pciexpresssettingdata
         *  
         *  Limitations:
         *      - Windows 10 1703/Windows Server 2016 or later is required for this method to work.
         *      - Changing host device name can cause counter effects to the Hyper-V system
         *      - WMI Filtering may affect the operation.
         */
        public void MountIntoVM(string vmName, string deviceId)
        {
            // 1. Connect to Hyper-V WMI
            Connect("root\\virtualization\\v2");

            uint outParams = 32768;
            string hostRes = string.Empty, vmRes = string.Empty;

            ManagementObject vmCurrentSetting = null;

            // 2. Create new instance for the vm setting
            ManagementObject setting = new ManagementClass(
                scope,
                new ManagementPath("Msvm_PciExpressSettingData"),
                null
            ).CreateInstance();

            // Get the service instance
            ManagementObject srv = new ManagementClass(
                scope,
                new ManagementPath("Msvm_VirtualSystemManagementService"),
                null
            ).GetInstances().Cast<ManagementObject>().FirstOrDefault();

            if (srv == null)
            {
                throw new ManagementException(
                    "MountIntoVM: Assignment service is either not running or crashed"
                );
            }

            // 3. From the PCIP device list, find if there is any device matches the PCI that has mounted
            //  From that, set the host resources (hostRes)
            using (ManagementObjectCollection pciExpressList = GetObjects("Msvm_PciExpress", "*"))
            {
                foreach (ManagementObject device in pciExpressList.Cast<ManagementObject>())
                {
                    // Get the device instance path (starts with "PCIP\")
                    string devInstancePath = device["DeviceInstancePath"]?.ToString();

                    if (devInstancePath == null)
                    {
                        device.Dispose();
                        continue;
                    }

                    // If matches, make sure the 
                    if (devInstancePath.Contains(deviceId.Replace("PCI\\", "PCIP\\")))
                    {
                        // Get the host name and device PnP
                        string hostName = device["SystemName"]?.ToString();
                        string deviceIdPnP = device["DeviceId"]?.ToString();

                        // Set the hostRes
                        /*
                         * Note it has a structure of this (in a single line, just seperate into lines so can see it easily):
                         *  \\{hostName}\root\virtualization\v2:
                         *  Msvm_PciExpress.CreationClassName="Msvm_PciExpress",
                         *  DeviceID="{deviceIdPnP}",
                         *  SystemCreationClassName="Msvm_ComputerSystem",
                         *  SystemName="{hostName}"
                         *  
                         *  where "hostName" is the device name, and "deviceIdPnP" is the device ID in GUIDv4
                         */
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
            }
            
            // If the hostRes is empty, throw new exception
            if (hostRes == string.Empty)
            {
                throw new ManagementException("MountIntoVM: Unable to get HostResource");
            }

            // 4. Get the VM setting to mount new device into the VM
            using (ManagementObjectCollection vmSettingDataList = GetObjects("Msvm_VirtualSystemSettingData", "*"))
            {
                foreach (ManagementObject vmSetting in vmSettingDataList.Cast<ManagementObject>())
                {
                    // Get the Caption and Instance ID
                    string vmCaption = vmSetting["Caption"]?.ToString();
                    string vmInstanceID = vmSetting["InstanceID"]?.ToString();

                    // Check if the vmCaption is null
                    if (vmCaption == null || vmInstanceID == null)
                    {
                        vmSetting.Dispose();
                        continue;
                    }

                    if (vmCaption.Equals("Virtual Machine Settings") && !vmInstanceID.Contains("Microsoft:Definition"))
                    {
                        // Get the VM name
                        string hostName = vmSetting["ElementName"]?.ToString();

                        if (hostName == null)
                        {
                            vmSetting.Dispose();
                            continue;
                        }

                        // If the VM name matches with the name wants to set the target, set the vmCurrentSetting and vmRes to it
                        if (hostName.Equals(vmName))
                        {
                            vmRes = $"{vmSetting["InstanceID"]}\\{Guid.NewGuid().ToString().ToUpper()}";
                            vmCurrentSetting = vmSetting;
                            vmSetting.Dispose();
                            break;
                        }
                    }
                }
            }
            
            if (vmRes == string.Empty)
            {
                throw new ManagementException("MountIntoVM: Unable to generate InstanceID");
            }

            if (vmCurrentSetting == null)
            {
                throw new ManagementException("MountIntoVM: Unable to get setting for VM");
            }

            // 5. Now try to mount the VM using vmCurrentSetting and setting
            try
            {
                // CIM_SettingData
                setting["Caption"] = "PCI Express Port";
                setting["Description"] = "Microsoft Virtual PCI Express Port Setting Data";
                setting["InstanceID"] = vmRes;
                setting["ElementName"] = "PCI Express Port";

                // CIM_ResourceAllocationSettingData
                setting["ResourceType"] = (ushort)32769;
                setting["OtherResourceType"] = null;
                setting["ResourceSubType"] = "Microsoft:Hyper-V:Virtual Pci Express Port";
                setting["PoolID"] = string.Empty;
                setting["ConsumerVisibility"] = (ushort)3;
                setting["HostResource"] = new string[] { hostRes };
                setting["AllocationUnits"] = "count";
                setting["VirtualQuantity"] = (ulong)1;
                setting["Reservation"] = (ulong)1;
                setting["Limit"] = (ulong)1;
                setting["Weight"] = (uint)0;
                setting["AutomaticAllocation"] = true;
                setting["AutomaticDeallocation"] = true;
                setting["Parent"] = null;
                setting["Connection"] = null;
                setting["Address"] = string.Empty;
                setting["MappingBehavior"] = null;
                setting["AddressOnParent"] = string.Empty;
                setting["VirtualQuantityUnits"] = "count";

                // Msvm_PciExpressSettingData
                setting["VirtualFunctions"] = new ushort[] { 0 };
                setting["VirtualSystemIdentifiers"] = new string[] { "{" + Guid.NewGuid() + "}" };

                // Call for the service
                outParams = (uint)srv.InvokeMethod("AddResourceSettings", new object[]
                {
                    vmCurrentSetting,
                    new string[] {
                        setting.GetText(TextFormat.WmiDtd20)
                    }
                });
            }
            finally
            {
                setting?.Dispose();
                srv?.Dispose();
            }

            // 6. Check for the out param. Returns 0 in case success, other numbers as failure, and failback as unknown
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
                    throw new ManagementException("MountIntoVM: Method Parameters Checked - Job Started");
                case 4097:
                    throw new ManagementException("MountIntoVM: The function may not be called or is reserved for vendor");
                default:
                    throw new ManagementException($"MountIntoVM: Unknown error ({outParams})");
            }
        }

        /*
         * Dismount PCIP device into the VM (Must be invoked when it is still in PCIP)
         * 
         * How does it work?
         *   1. It connects to the Hyper-V namespace (root\virtualization\v2)
         *   2. From the provided variable deviceId, try to get the PCIP device that match deviceID 
         *      in Msvm_PciExpressSettingData
         *   3. Call the service Msvm_PciExpressSettingData to dismount the device from the VM
         *   4. From the out parameter, check the return value and compare it into Success or 
         *      Failure state
         *   
         *  Note:
         *      - For each value of the get ManagementObject, it is required to check for null value
         *      - No matter how the service operate with the parameter, any WMI object must be disposed
         *      - The function must be invoked from the PCI, that means the when the device is still with
         *      
         *  References: 
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-virtualsystemmanagementservice
         *      - https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-pciexpresssettingdata
         *  
         *  Limitations:
         *      - Windows 10 1703/Windows Server 2016 or later is required for this method to work.
         *      - WMI Filtering may affect the operation.
         */
        public void DismountFromVM(string deviceId)
        {
            // 1. Connect to Hyper-V WMI
            Connect("root\\virtualization\\v2");
            uint outParams = 32768;

            ManagementObject[] deviceObj = new ManagementObject[1];

            // Get the service
            ManagementObject srv = new ManagementClass(scope, new ManagementPath("Msvm_VirtualSystemManagementService"), null).GetInstances().Cast<ManagementObject>().FirstOrDefault();

            if (srv == null)
            {
                throw new ManagementException("DismountFromVM: Assignment service is either not running or crashed");
            }

            // 2. Get the device instance from deviceId
            using (ManagementObjectCollection pciExpressSettingList = GetObjects("Msvm_PciExpressSettingData", "*"))
            {
                foreach (ManagementObject obj in pciExpressSettingList.Cast<ManagementObject>())
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
            }

            if (deviceObj[0] == null)
            {
                throw new ManagementException("DismountFromVM: Error while finding the device ID");
            }

            // 3. Call the service
            try
            {
                outParams = (uint)srv.InvokeMethod("RemoveResourceSettings", new object[]
                {
                    deviceObj
                });
            }
            finally
            {
                srv?.Dispose();
            }

            // 4. Check for the out param.Returns 0 in case success, other numbers as failure, and failback as unknown
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
                    throw new ManagementException("DismountFromVM: Method Parameters Checked - Job Started");
                default:
                    throw new ManagementException($"DismountFromVM: Unknown error ({outParams})");
            }
        }
    }
}