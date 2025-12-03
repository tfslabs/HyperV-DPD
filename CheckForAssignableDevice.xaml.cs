using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Threading.Tasks;
using System.Windows;
using TheFlightSims.HyperVDPD.DefaultUI;
using TheFlightSims.HyperVDPD.WMIProperties;

/*
 * Primary namespace for HyperV-DPD application
 *  It contains the main window and all related methods for the core application
 */
namespace TheFlightSims.HyperVDPD
{
    /*
     * Check for assignable device window
     * It is used to check for devices that are assignable for PCIP
     */
    public partial class CheckForAssignableDevice : Window
    {
        ////////////////////////////////////////////////////////////////
        /// Global Properties and Constructors Region
        ///     This region contains global properties and constructors 
        ///     for the MachineMethods class.
        ////////////////////////////////////////////////////////////////

        /*
         * Global properties
         *  machine: Instance of MachineMethods class to interact with WMI
         */
        protected MachineMethods machine;

        /*
         * Constructor of the class
         */
        public CheckForAssignableDevice(MachineMethods machine)
        {
            this.machine = machine;
            InitializeComponent();
        }

        ////////////////////////////////////////////////////////////////
        /// User Action Methods Region
        ///     This region contains methods that handle user actions.
        ///     For example, button clicks, changes in order.
        ////////////////////////////////////////////////////////////////

        /*
         * Actions for "Check" button
         */
        private async void Check_Button(object sender, RoutedEventArgs e)
        {
            await CheckDevices();
        }

        /*
         * Actions for "Cancel" button
         */
        private void Cancel_Button(object sender, RoutedEventArgs e)
        {
            Close();
        }

        ////////////////////////////////////////////////////////////////
        /// Non-User Action Methods Region
        /// 
        /// This region contains methods that do not handle user actions.
        /// 
        /// Think about this is the back-end section.
        /// It should not be in a seperated class, because it directly interacts with the UI elements.
        ////////////////////////////////////////////////////////////////

        /*
         * Task for Check devices
         */
        private async Task CheckDevices()
        {
            // Clear the device list
            AssignableDevice_ListView.Items.Clear();

            /*
             * Check for assignable devices
             * 
             * How does it work?
             *    1. It looks up for device data in
             *      - Win32_PnPEntity, to get its Hardware ID, Hardware Friendly Name (Caption), and current status. 
             *          Construct all data into a list, follows this structure of a tuple:
             *              (string deviceID, string devCaption, devStatus)
             *      
             *      - Win32_PnPDevice, to find if the device is actual device
             *          Construct all data in a hash set, where the devices are Hardware ID unique.
             *      
             *      - Win32_PNPAllocatedResource, to find the starting address of the Hardware ID. 
             *          Note: A device may have multiple starting memory
             *          Construct into a list, follows this structure:
             *              (UInt64 startingAddress, string deviceID)
             *      
             *      - Win32_DeviceMemoryAddress, to look up for possible address range, and calculate the required memory
             *          Construct into a hash map, follows this structure:
             *              <TKey: startingAddress, TValue: endingAddress>
             *
             *    2. For each device data, put into the boxes. Converting values from MOF structure into processable data can be seen below.
             *    3. From the data structure, try to validate if a device from the Win32_PnPEntity matches the following condition
             *          - The device is PCI device (required)
             *          - The device exists in Win32_PnPDevice (required)
             *          - If possible, check for the memory gap. Note that if the device is disabled, the gap cannot be calculated
             *       If any the required condition does not meet, notify user for the "unable to assignment" and display for memory gap as "N/A"
             *       If unable to calculate the gap, just display "N/A"
             *  
             * Note:
             *      - All device ID must be normalized with a single "\" and not "\\"
             *      - Normally, the WMI object's property is formed into multivalue, and value seperation with validation is required.
             *      - Check for null, empty, or default (fallback) value in each case you get the value from the WMI.
             *      - Disposing the resouce after using it
             * 
             * References:
             *      - https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/win32-pnpentity
             *      - https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/win32-pnpdevice
             *      - https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/win32-pnpallocatedresource
             *      - https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/win32-devicememoryaddress
             * 
             * Limitation:
             *      Each resource has its own limitations. See below
             */
            try
            {
                await Task.Run(() =>
                {
                    machine.Connect("root\\cimv2");

                    // API calls to get all PnP entities
                    ManagementObjectCollection pnpDeviceEntity = machine.GetObjects("Win32_PnPEntity", "DeviceID, Caption, Status");
                    ManagementObjectCollection pnpDevice = machine.GetObjects("Win32_PnPDevice", "SystemElement");
                    ManagementObjectCollection pnpDeviceAllocatedResource = machine.GetObjects("Win32_PNPAllocatedResource", "Antecedent, Dependent");
                    ManagementObjectCollection pnpDeviceMemorySpace = machine.GetObjects("Win32_DeviceMemoryAddress", "StartingAddress, EndingAddress");

                    /*
                     * Device list on the computer
                     * 
                     * pnpDeviceEntityMap - List
                     * 
                     * Data structure
                     *  deviceId (string): The ID of the device on the system. Example: PCI\VEN_10DE&DEV_2520&SUBSYS_12F21462&REV_A1
                     *  devCaption (string): Actual device name. Example: NVIDIA GeForce RTX 3060 Laptop GPU
                     *  devStatus (string): The status of the device. See the Win32_PnPEntity for detailed status info
                     */
                    List<(string deviceId, string devCaption, string devStatus)> pnpDeviceEntityMap = new List<(string, string, string)>();
                    foreach (ManagementObject device in pnpDeviceEntity.Cast<ManagementObject>())
                    {
                        string deviceId = device["DeviceID"]?.ToString();
                        string devCaption = device["Caption"]?.ToString();
                        string devStatus = device["Status"]?.ToString();

                        if (!string.IsNullOrEmpty(deviceId) && !string.IsNullOrEmpty(devCaption) && !string.IsNullOrEmpty(devStatus))
                        {
                            // Don't forget to normalize the device id
                            pnpDeviceEntityMap.Add(
                                (deviceId.Replace("\\\\", "\\").Trim(), devCaption, devStatus)
                            );
                        }

                        // Dispose after using it
                        device.Dispose();
                    }

                    /*
                     * PnP device on the computer
                     * 
                     * pnpDeviceSet - Hash Set
                     * 
                     * Data structure
                     *  actualPnpDevString (string): the device id exist on the class
                     *  
                     * How does the WMI property contribute for data
                     *  SystemElement: \\<Computer Name>\root\cimv2:Win32_PnPEntity.DeviceID="<Un-normalized device ID data>"
                     * 
                     */
                    HashSet<string> pnpDeviceSet = new HashSet<string>();
                    foreach (ManagementObject device in pnpDevice.Cast<ManagementObject>())
                    {
                        string PnpDevString = device["SystemElement"]?.ToString();
                        
                        if (!string.IsNullOrEmpty(PnpDevString))
                        {
                            // Seperate between the path and pointer
                            string[] SystemElementArr = PnpDevString.Split(':');
                            
                            // In case the pointer does not exist (Really?)
                            if (SystemElementArr.Length < 2)
                            {
                                // Dispose the device and go on
                                device.Dispose();
                                continue;
                            }

                            // The array[1] is the pointer.
                            string actualPnpDevString = SystemElementArr[1].Replace("Win32_PnPEntity.DeviceID=", "").Trim('"');
                            
                            // Add into the hash set. Don't forget to normalize
                            pnpDeviceSet.Add(actualPnpDevString.Replace("\\\\", "\\"));
                        }

                        // Dispose the device instance
                        device.Dispose();
                    }

                    /*
                     * PnP Device Allocation Map
                     * 
                     * pnpDeviceAllocatedResourceMap - List
                     * 
                     * Data structure
                     *  startingAddress (UInt64): starting address
                     *  deviceID (string): device ID
                     *  
                     * How does the WMI property contribute for data
                     *  Antecedent: \\<Computer Name>\root\cimv2:Win32_DeviceMemoryAddress.StartingAddress="<Starting Address in UInt64>"
                     *  Dependent: \\<Computer Name>\root\cimv2:Win32_PnPEntity.DeviceID="<Un-normalized device ID data>"
                     */
                    List<(UInt64 startingAddress, string deviceID)> pnpDeviceAllocatedResourceMap = new List<(UInt64 startingAddress, string deviceID)>();
                    foreach (ManagementObject resource in pnpDeviceAllocatedResource.Cast<ManagementObject>())
                    {
                        string antecedentRes = resource["Antecedent"]?.ToString();
                        string dependentRes = resource["Dependent"]?.ToString();
                        
                        if (!string.IsNullOrEmpty(antecedentRes) && !string.IsNullOrEmpty(dependentRes))
                        {
                            // Seperate between the path and pointer
                            string[] AntecedentArr = antecedentRes.Split(':');
                            string[] DependentArr = dependentRes.Split(':');

                            // In case the pointer does not exist (Really?)
                            if (AntecedentArr.Length < 2 || DependentArr.Length < 2)
                            {
                                // Dispose the device and go on
                                resource.Dispose();
                                continue;
                            }

                            // The array[1] is the pointer.
                            UInt64 startingAddress = UInt64.Parse(AntecedentArr[1].Split('=')[1].Trim('"'));
                            string deviceID = DependentArr[1].Replace("Win32_PnPEntity.DeviceID=", "").Trim('"');

                            // Add into the hash set. Don't forget to normalize
                            pnpDeviceAllocatedResourceMap.Add(
                                (startingAddress, deviceID.Replace("\\\\", "\\"))
                            );
                        }

                        resource.Dispose();
                    }

                    /*
                     * PnP Device Memory Space Map
                     * 
                     * pnpDeviceMemorySpaceMap - Dictionary/Hash Map
                     * 
                     * Data structure
                     *  startingAddress (UInt64): Starting address (unique)
                     *  endingAddress (UInt64): Ending address
                     */
                    Dictionary<UInt64, UInt64> pnpDeviceMemorySpaceMap = new Dictionary<UInt64, UInt64>();
                    foreach (ManagementObject memorySpaces in pnpDeviceMemorySpace.Cast<ManagementObject>())
                    {
                        string startingAddress = memorySpaces["StartingAddress"]?.ToString();
                        string endingAddress = memorySpaces["EndingAddress"]?.ToString();

                        if (!string.IsNullOrEmpty(startingAddress) && !string.IsNullOrEmpty(endingAddress))
                        {
                            pnpDeviceMemorySpaceMap.Add(
                                UInt64.Parse(startingAddress),
                                UInt64.Parse(endingAddress)
                            );
                        }

                        memorySpaces.Dispose();
                    }

                    // Dispose all resources, since the data has been processed
                    pnpDeviceEntity.Dispose();
                    pnpDevice.Dispose();
                    pnpDeviceAllocatedResource.Dispose();
                    pnpDeviceMemorySpace.Dispose();

                    foreach ((string id, string caption, string status) in pnpDeviceEntityMap)
                    {
                        bool isPciDevice = true, isActualDevice = true, isEnabled = true;
                        string deviceNote = string.Empty;
                        UInt64 memoryGap = 0;

                        if (!id.StartsWith("PCI\\"))
                        {
                            isPciDevice = false;
                            deviceNote += ((deviceNote.Length == 0) ? "" : "\n") + "The device is not a PCI.";
                        }

                        if (!pnpDeviceSet.Contains(id))
                        {
                            isActualDevice = false;
                            deviceNote += ((deviceNote.Length == 0) ? "" : "\n") +
                                "The PCI device is not support either Express Endpoint, Embedded Endpoint, or Legacy Express Endpoint.";
                        }

                        if (!string.Equals(status, "OK", StringComparison.OrdinalIgnoreCase))
                        {
                            isEnabled = false;
                            deviceNote += ((deviceNote.Length == 0) ? "" : "\n") +
                                "Device is disabled. Please re-enable the device to let the program check the memory gap.";
                        }

                        if (isPciDevice && isActualDevice && isEnabled)
                        {
                            foreach ((UInt64 startMemory, string deviceId) in pnpDeviceAllocatedResourceMap)
                            {
                                if (id.Equals(deviceId))
                                {
                                    if (pnpDeviceMemorySpaceMap.ContainsKey(startMemory))
                                    {
                                        UInt64 endAddr = pnpDeviceMemorySpaceMap[startMemory];
                                        memoryGap += (endAddr - startMemory + 1048575UL) / 1048576UL;
                                    }
                                }
                            }
                        }

                        Dispatcher.Invoke(() =>
                        {
                            _ = AssignableDevice_ListView.Items.Add(new
                            {
                                DeviceId = id,
                                DeviceName = caption,
                                DeviceGap = (isPciDevice && isActualDevice && isEnabled) ? $"{memoryGap} MB" : "N/A",
                                DeviceNote = deviceNote
                            });
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                (new ExceptionView()).HandleException(ex);
            }
        }
    }
}
