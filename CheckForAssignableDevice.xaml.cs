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

            AssignableDevice_ListView.Items.Clear();

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

                    List<(string deviceId, string devCaption, string devStatus)> pnpDeviceEntityMap = new List<(string, string, string)>();
                    foreach (ManagementObject device in pnpDeviceEntity.Cast<ManagementObject>())
                    {
                        string deviceId = device["DeviceID"]?.ToString();
                        string devCaption = device["Caption"]?.ToString();
                        string devStatus = device["Status"]?.ToString();

                        if (!string.IsNullOrEmpty(deviceId) && !string.IsNullOrEmpty(devCaption) && !string.IsNullOrEmpty(devStatus))
                        {
                            pnpDeviceEntityMap.Add(
                                (deviceId.Replace("\\\\", "\\").Trim(), devCaption, devStatus)
                            );
                        }

                        device.Dispose();
                    }

                    HashSet<string> pnpDeviceSet = new HashSet<string>();
                    foreach (ManagementObject device in pnpDevice.Cast<ManagementObject>())
                    {
                        string PnpDevString = device["SystemElement"]?.ToString();

                        if (!string.IsNullOrEmpty(PnpDevString))
                        {
                            string actualPnpDevString = PnpDevString.Split('=')[1].Replace("\"", "").Replace("\\\\", "\\").Trim();
                            pnpDeviceSet.Add(actualPnpDevString);
                        }

                        device.Dispose();
                    }

                    List<(UInt64 startingMemory, string deviceID)> pnpDeviceAllocatedResourceMap = new List<(UInt64 startingMemory, string deviceID)>();
                    foreach (ManagementObject resource in pnpDeviceAllocatedResource.Cast<ManagementObject>())
                    {
                        string antecedentRes = resource["Antecedent"]?.ToString();
                        string dependentRes = resource["Dependent"]?.ToString();

                        if (!string.IsNullOrEmpty(antecedentRes) && !string.IsNullOrEmpty(dependentRes))
                        {
                            UInt64 startingMemory = UInt64.Parse(antecedentRes.Split('=')[1].Replace("\"", ""));
                            string deviceID = dependentRes.Split('=')[1].Replace("\"", "").Replace("\\\\", "\\").Trim();

                            pnpDeviceAllocatedResourceMap.Add((
                                    startingMemory,
                                    deviceID
                            ));
                        }

                        resource.Dispose();
                    }

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

                    
                    /*
                     * 
                     */
                    foreach ((string id, string caption, string status) in pnpDeviceEntityMap)
                    {
                        if (id.StartsWith("PCI\\"))
                        {
                            bool isAssignable = true;
                            UInt64 memoryGap = 0;
                            string deviceNote = string.Empty;

                            if (!status.ToLower().Equals("ok"))
                            {
                                isAssignable = false;
                                deviceNote += ((deviceNote.Length == 0) ? "" : "\n") + "Device is disabled. Please re-enable the device to let the program check the memory gap.";
                            }

                            if (!pnpDeviceSet.Contains(id))
                            {
                                isAssignable = false;
                                deviceNote += ((deviceNote.Length == 0) ? "" : "\n") + "The PCI device is not support either Express Endpoint, Embedded Endpoint, or Legacy Express Endpoint.";
                            }

                            if (isAssignable)
                            {
                                foreach ((UInt64 startMemory, string deviceId) deviceResource in pnpDeviceAllocatedResourceMap)
                                {
                                    if (id.Equals(deviceResource.deviceId))
                                    {
                                        if (pnpDeviceMemorySpaceMap.ContainsKey(deviceResource.startMemory))
                                        {
                                            UInt64 endAddr = pnpDeviceMemorySpaceMap[deviceResource.startMemory];
                                            memoryGap += (endAddr - deviceResource.startMemory + 1048575) / 1048576;
                                        }
                                    }
                                }
                            }
                            
                            if (memoryGap == 0)
                            {
                                deviceNote += ((deviceNote.Length == 0) ? "" : "\n") + "Unable to calculate gap memory.";
                            }

                            Dispatcher.Invoke(() =>
                            {
                                _ = AssignableDevice_ListView.Items.Add(new
                                {
                                    DeviceId = id,
                                    DeviceName = caption,
                                    DeviceGap = (!isAssignable) ? "N/A" : $"{memoryGap} MB",
                                    DeviceNote = deviceNote
                                });
                            });
                        }
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
