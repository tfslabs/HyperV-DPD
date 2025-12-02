using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Security;
using System.Text.RegularExpressions;
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
     * Main Window class
     * It is used to display the main UI of the application
     */
    public partial class MainWindow : Window
    {
        ////////////////////////////////////////////////////////////////
        /// Global Properties and Constructors Region
        ///     This region contains global properties and constructors 
        ///     for the MachineMethods class.
        ////////////////////////////////////////////////////////////////

        /*
         *  Global properties
         *  machine: Instance of MachineMethods class to interact with Hyper-V
         *  vmObjects: Dictionary to store VM information including name, status, and assigned devices
         */
        private MachineMethods machine;
        private readonly Dictionary<string, (string vmName, string vmStatus, List<(string instanceId, string devInstancePath)> devices)> vmObjects;

        // Will do migrate later
        private readonly Dictionary<string, (string vmName, string vmStatus)> vmStatus;
        private readonly Dictionary<string, List<(string instanceId, string devInstancePath)>> vmDevices;

        // Default values of VM status mapping
        public static readonly Dictionary<ushort, string> vmStatusMap = new Dictionary<ushort, string>
        {
            {0,  "Unknown" },
            {1,  "Other" },
            {2,  "Running" },
            {3,  "Stopped" },
            {4,  "Shutting down" },
            {5,  "Not applicable" },
            {6,  "Enabled but Offline" },
            {7,  "In Test" },
            {8,  "Degraded" },
            {9,  "Quiesce" },
            {10, "Starting" }
        };

        /*
         * Constructor of the MainWindow class
         * Initializes the components and sets up event handlers
         */
        public MainWindow()
        {
            InitializeComponent();

            // Initialize global properties
            machine = new MachineMethods();
            vmObjects = new Dictionary<string, (string vmName, string vmStatus, List<(string instanceId, string devInstancePath)> devices)>();

            // Will do migrate later
            vmStatus = new Dictionary<string, (string vmName, string vmStatus)>();
            vmDevices = new Dictionary<string, List<(string instanceId, string devInstancePath)>>();

            // Set up the sync loading event handler. See the method below for further info.
            Loaded += MainWindow_Loaded;
        }

        /*
         * Main Windows async loading method.
         * It is used prevent long loading times on the UI thread
         */
        public async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await ReinitializeConnection();
        }

        ////////////////////////////////////////////////////////////////
        /// User Action Methods Region
        ///     This region contains methods that handle user actions.
        ///     For example, button clicks, changes in order.
        ////////////////////////////////////////////////////////////////

        /*
         * Actions for button "Connect to Another Computer"
         */
        private async void ConnectToAnotherComputer_Button(object sender, RoutedEventArgs e)
        {
            // Update status bar to show connecting status
            StatusBarChangeBehaviour(true, "Connecting");

            // Show the connection form to get user credentials
            (string computerName, string username, SecureString password) userCredential = (new ConnectForm()).ReturnValue();

            // If the user provided a computer name, create a new MachineMethods instance with the provided credentials
            if (userCredential.computerName != "")
            {
                // Decide whether to connect locally or remotely based on the computer name
                machine = (string.Equals(userCredential.computerName, "localhost", StringComparison.OrdinalIgnoreCase)) ? new MachineMethods() : new MachineMethods(userCredential);
                // Reinitialize the connection to the specified computer
                await ReinitializeConnection();
            }
        }

        /*
         * Actions for button "Quit"
         */
        private void QuitMainWindow_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        /*
         * Actions for button "Add Device"
         */
        private async void AddDevice_Click(object sender, RoutedEventArgs e)
        {
            // Check if the current VM selection is null (nothing is selected)
            if (VMList.SelectedItem != null)
            {
                // Change the status bar
                StatusBarChangeBehaviour(true, "Adding device");

                // Get the device ID from the form Add device
                string deviceId = (new AddDevice(machine)).GetDeviceId();

                // Get the vmName from the selection
                string vmName = VMList.SelectedItem.GetType().GetProperty("VMName").GetValue(VMList.SelectedItem, null)?.ToString();

                // Check both if null then reject
                if (deviceId != null && vmName != null)
                {
                    // Try to mount the device into the virtual machine
                    // If failed, try unmount and re-enable
                    try
                    {
                        /* Run for this following task:
                         * 1. Try to disable the device from the PnP
                         * 2. Then, mount it from PCI module into PCIP module of Hyper-V
                         * 3. Lastly, mount the device into the VM and let the VM do it best
                         */
                        await Task.Run(() =>
                        {
                            machine.ChangePnpDeviceBehaviour(deviceId, isDisable: true);
                            machine.MountPnPDeviceToPcip(deviceId);
                            machine.MountIntoVM(vmName, deviceId);
                        });

                        // Try refreshing the VM status
                        await RefreshVMs();
                    }
                    catch (Exception ex)
                    {
                        // In case any problem occur while mounting,
                        // try dismount from PCIP the device and re-enable

                        // Display for error
                        (new ExceptionView()).HandleException(ex);
                        StatusBarChangeBehaviour(false, "Error, Re-enable device");

                        // Try re-enable device
                        // If still fail, leave it alone
                        try
                        {
                            await Task.Run(() =>
                            {
                                machine.ChangePnpDeviceBehaviour(deviceId, isDisable: false);
                            });
                        }
                        catch (Exception exp)
                        {
                            (new ExceptionView()).HandleException(exp);
                            StatusBarChangeBehaviour(false, "Re-enable device failed");
                        }
                    }
                }
                else
                {
                    // If nothing is selected, check for device
                    StatusBarChangeBehaviour(false, "No Device Added");
                }
            }
            else
            {
                // If there is no VM selected on the list
                _ = MessageBox.Show(
                    "Please select a virtual machine on the list",
                    "Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                StatusBarChangeBehaviour(false, "No Device Added");
            }
        }

        /*
         * Actions for button "Change Memory Location"
         */
        private async void ChangeMemLocation_Click(object sender, RoutedEventArgs e)
        {
            // Check if the current VM selection is null (nothing is selected)
            if (VMList.SelectedItem != null)
            {
                // Try to change memory allocation for the target VM
                try
                {
                    // Get the vmName from the selection
                    string vmName = VMList.SelectedItem.GetType().GetProperty("VMName").GetValue(VMList.SelectedItem, null).ToString();

                    // Get the memory range from the form Change Memory Space
                    (ulong lowMem, ulong highMem) = (new ChangeMemorySpace(vmName)).ReturnValue();

                    // Validate both values
                    if (lowMem != 0 && highMem != 0)
                    {
                        // If valid, change the memory allocation
                        await Task.Run(() =>
                        {
                            machine.ChangeMemAllocate(vmName, lowMem, highMem);
                        });

                        // Try refreshing the VM status
                        await RefreshVMs();
                    }
                }
                catch (Exception ex)
                {
                    // In case any problem occur while changing memory allocation,
                    (new ExceptionView()).HandleException(ex);
                    StatusBarChangeBehaviour(false, "Error");
                }
            }
            else
            {
                // If there is no VM selected on the list
                _ = MessageBox.Show(
                    "Please select a virtual machine on the list",
                    "Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                StatusBarChangeBehaviour(false, "No virtual machine Selected");
            }
        }

        /*
         * Remove device from VM button action
         */
        private async void RemDevice_Click(object sender, RoutedEventArgs e)
        {
            // Check if there is a VM selected and a device selected
            if (VMList.SelectedItem != null && DevicePerVMList.SelectedItem != null)
            {
                // Try to remove the device from the VM
                try
                {
                    // Try to get the device ID and device path from the selected item
                    string deviceId = DevicePerVMList.SelectedItem.GetType().GetProperty("DeviceID").GetValue(DevicePerVMList.SelectedItem, null)?.ToString();
                    string devicePath = DevicePerVMList.SelectedItem.GetType().GetProperty("DevicePath").GetValue(DevicePerVMList.SelectedItem, null)?.ToString();

                    StatusBarChangeBehaviour(true, "Removing device");

                    // Check both if null then reject
                    if (deviceId != null && devicePath != null)
                    {
                        // Run for this following task:
                        // 1. Dismount the device from the VM
                        // 2. Dismount the device from the PCIP module
                        // 3. Re-enable the device from PnP
                        await Task.Run(() =>
                        {
                            machine.DismountFromVM(deviceId);
                            machine.DismountPnPDeviceFromPcip(devicePath);
                            machine.ChangePnpDeviceBehaviour(devicePath.Replace("PCIP\\", "PCI\\"), isDisable: false);
                        });

                        // Refresh the VM list
                        await RefreshVMs();
                    }
                }
                catch (Exception ex)
                {
                    // In case any problem occur while removing the device,
                    (new ExceptionView()).HandleException(ex);
                    StatusBarChangeBehaviour(false, "Error");
                }
            }
            else
            {
                // If there is no VM or device selected on the list
                _ = MessageBox.Show(
                    "Please select a device from the virtual machine list!",
                    "Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );

                StatusBarChangeBehaviour(false, "No Device Selected");
            }
        }

        /*
         * Copy device address button action
         */
        private void CopyDevAddress_Click(object sender, RoutedEventArgs e)
        {
            // Check if there is a VM selected and a device selected
            if (VMList.SelectedItem != null && DevicePerVMList.SelectedItem != null)
            {
                // Try to get the device path from the selected item
                string devicePath = DevicePerVMList.SelectedItem.GetType().GetProperty("DevicePath").GetValue(DevicePerVMList.SelectedItem, null)?.ToString();

                // If not null, copy to clipboard
                if (devicePath != null)
                {
                    // Copy to clipboard
                    Clipboard.SetText(devicePath);
                    _ = MessageBox.Show(
                        $"Copied device Path {devicePath} into clipboard",
                        "Info",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information
                    );
                }
            }
            else
            {
                // If there is no VM or device selected on the list
                _ = MessageBox.Show(
                    "Please select a virtual machine and a device!",
                    "Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                StatusBarChangeBehaviour(false, "No Device Selected");
            }
        }

        /*
         * Actions for button "Hyper-V Service Status"
         */
        private void HyperVServiceStatus_Click(object sender, RoutedEventArgs e)
        {
            _ = (new HyperVStatus(machine)).ShowDialog();
        }

        /*
         * Actions for button "Check for Assignable Device"
         */
        private void AssignableDevice_Click(object sender, RoutedEventArgs e)
        {
            _ = (new CheckForAssignableDevice(machine)).ShowDialog();
        }

        /*
         * Actions for button "Remove All Devices"
         */
        private async void RemAllDevice_Click(object sender, RoutedEventArgs e)
        {
            // Try to remove all assigned devices from all VMs
            try
            {
                // Confirm with the user first
                MessageBoxResult isRemoveAll = MessageBox.Show(
                            $"Do you want to remove all assigned devices? (This option should only be used in case you accidently remove a virtual machine without unmounting it first)",
                            "Remove all devices",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question
                );

                // If yes, proceed to remove all devices
                if (isRemoveAll == MessageBoxResult.Yes)
                {
                    StatusBarChangeBehaviour(true, "In action");

                    // Run through all VMs and remove all devices
                    await Task.Run(() =>
                    {
                        // Connect to Hyper-V WMI namespace
                        machine.Connect("root\\virtualization\\v2");

                        // Iterate through all PCI Express Setting Data objects
                        using (ManagementObjectCollection devSettings = machine.GetObjects("Msvm_PciExpressSettingData", "InstanceID"))
                        {
                            // If there are any devices assigned, dismount them from their respective VMs
                            if (devSettings != null)
                            {
                                // Iterate through each device setting
                                foreach (ManagementObject deviceid in devSettings.Cast<ManagementObject>())
                                {
                                    //  Get the InstanceID of the device
                                    string devInstanceId = deviceid["InstanceID"]?.ToString();

                                    // If InstanceID is null, skip to the next device
                                    if (devInstanceId == null)
                                    {
                                        deviceid.Dispose();
                                        continue;
                                    }

                                    // Dismount the device from its VM if it's not a Microsoft Definition device
                                    if (!devInstanceId.Contains("Microsoft:Definition"))
                                    {
                                        machine.DismountFromVM(devInstanceId);
                                    }

                                    // Dispose of the ManagementObject to free resources
                                    deviceid.Dispose();
                                }
                            }
                        }

                        // Iterate through all PCI Express devices in the PCIP namespace
                        using (ManagementObjectCollection devMount = machine.GetObjects("Msvm_PciExpress", "DeviceInstancePath"))
                        {
                            // If there are any devices mounted in PCIP, dismount them and re-enable in PnP
                            if (devMount != null)
                            {
                                // Iterate through each mounted device
                                foreach (ManagementObject devInstancePath in devMount.Cast<ManagementObject>())
                                {
                                    // Get the DeviceInstancePath of the device
                                    string devInsPath = devInstancePath["DeviceInstancePath"]?.ToString();

                                    // If DeviceInstancePath is not null, dismount and re-enable the device
                                    if (devInsPath != null)
                                    {
                                        // Dismount the device from PCIP
                                        machine.DismountPnPDeviceFromPcip(devInsPath);

                                        /*
                                         * Note: When re-enabling the device, we need to replace "PCIP\" with "PCI\"
                                         * Since the device was originally a PCI device before being assigned to Hyper-V
                                         */
                                        machine.ChangePnpDeviceBehaviour(devInsPath.Replace("PCIP\\", "PCI\\"), isDisable: false);

                                        // Dispose of the ManagementObject to free resources
                                        devInstancePath.Dispose();
                                    }
                                    else
                                    {
                                        // Dispose of the ManagementObject to free resources
                                        devInstancePath.Dispose();
                                        continue;
                                    }
                                }
                            }
                        }
                    });

                    // Refresh the VM list to reflect changes
                    await RefreshVMs();
                }
                else
                {
                    // If user chose not to remove all devices
                    StatusBarChangeBehaviour(false, "No Action Performed");
                }
            }
            catch (Exception ex)
            {
                // In case any problem occur while removing all devices,
                // display for error
                (new ExceptionView()).HandleException(ex);
                StatusBarChangeBehaviour(false, "Error");
            }
        }

        /*
         * Actions for button "Refresh"
         */
        private async void HyperVRefresh_Click(object sender, RoutedEventArgs e)
        {
            // Refresh the VM list
            await RefreshVMs();
        }

        /*
         * Actions for button "About"
         */
        private void AboutBox_Click(object sender, RoutedEventArgs e)
        {
            _ = (new About()).ShowDialog();
        }

        /*
         * Actions for button "Change Cache Types"
         */
        private async void ChangeCacheTypes_Click(object sender, RoutedEventArgs e)
        {
            // Check if the current VM selection is null (nothing is selected)
            if (VMList.SelectedItem != null)
            {
                // Try to change cache type for the target VM
                try
                {
                    // Get the vmName from the selection
                    string vmName = VMList.SelectedItem.GetType().GetProperty("VMName").GetValue(VMList.SelectedItem, null)?.ToString();

                    // Confirm with the user first
                    MessageBoxResult isEnableMemCache = MessageBox.Show(
                        $"Do you want to enable guest control cache type for {vmName}? (Yes to enable, No to disable, Cancel to leave it as is)",
                        "Enable Guest Control Cache",
                        MessageBoxButton.YesNoCancel,
                        MessageBoxImage.Question
                    );
                    
                    // If yes, proceed to enable memory cache
                    if (isEnableMemCache == MessageBoxResult.Yes)
                    {
                        // Enable guest control cache type
                        await Task.Run(() =>
                        {
                            machine.ChangeGuestCacheType(vmName, true);
                        });

                        // Notify user of success
                        _ = MessageBox.Show(
                            $"Enabled guest control cache type successfully for {vmName}",
                            "Info",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information
                        );
                    }
                    else if (isEnableMemCache == MessageBoxResult.No)
                    {
                        // Disable guest control cache type
                        await Task.Run(() =>
                        {
                            machine.ChangeGuestCacheType(vmName, false);
                        });

                        // Notify user of success
                        _ = MessageBox.Show(
                                $"Disabled guest control cache type successfully for {vmName}",
                                "Info",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information
                        );
                    }

                    // Try refreshing the VM status
                    await RefreshVMs();
                }
                catch (Exception ex)
                {
                    // In case any problem occur while changing cache type,
                    // display for error
                    (new ExceptionView()).HandleException(ex);
                    StatusBarChangeBehaviour(false, "Error");
                }
            }
            else
            {
                // If there is no VM selected on the list
                _ = MessageBox.Show(
                    "Please select a virtual machine to add a device",
                    "Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
            }
        }

        /*
         * Actions for selecting a VM from the list
         */
        private async void VMList_Select(object sender, RoutedEventArgs e)
        {
            // When a VM is selected, update the device list for that VM
            if (VMList.SelectedItem != null)
            {
                // Clear the current device list
                DevicePerVMList.Items.Clear();
                // Get the VM ID from the selected item
                string vmId = VMList.SelectedItem.GetType().GetProperty("VMId").GetValue(VMList.SelectedItem, null)?.ToString();

                // If VM ID is not null, populate the device list for that VM
                if (vmId != null)
                {
                    // Populate the device list asynchronously
                    await Task.Run(() =>
                    {
                        // Find the VM object in the dictionary
                        foreach (KeyValuePair<string, (string vmName, string vmStatus, List<(string instanceId, string devInstancePath)> devices)> vmObject in vmObjects)
                        {
                            // If the VM ID matches, add its devices to the device list
                            if (vmObject.Key.Equals(vmId))
                            {
                                // Add each device to the device list
                                foreach ((string instanceId, string devInstancePath) in vmObject.Value.devices)
                                {
                                    // Use Dispatcher to update the UI thread
                                    Dispatcher.Invoke(() =>
                                    {
                                        // Add the device to the DevicePerVMList
                                        _ = DevicePerVMList.Items.Add(new
                                        {
                                            DeviceID = instanceId,
                                            DevicePath = devInstancePath
                                        });
                                    });
                                }
                                // Exit the loop once the matching VM is found
                                break;
                            }
                        }
                    });
                }
            }
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
         * Reinitialize connection to the Hyper-V host
         */
        private async Task ReinitializeConnection()
        {
            // Flags to check system compatibility
            bool isHyperVDisabled = true, isOSTooOld = true, isOSNotServer = true;
            
            StatusBarChangeBehaviour(true, "Connecting");

            // Try to connect to the Hyper-V host and check system compatibility
            try
            {
                // Run the connection and compatibility check asynchronously
                await Task.Run(() =>
                {
                    // Connect to the CIMV2 namespace to check OS information
                    machine.Connect("root\\cimv2");
                    foreach (ManagementObject osInfo in machine.GetObjects("Win32_OperatingSystem", "BuildNumber, ProductType").Cast<ManagementObject>())
                    {
                        // Check OS build number for DDA support
                        string osInfoBuildNumber = osInfo["BuildNumber"]?.ToString() ?? "";
                        if (int.TryParse(osInfoBuildNumber, out int buildNumber))
                        {
                            // DDA is supported from Windows 10 build 14393 and later
                            if (buildNumber >= 14393)
                            {
                                isOSTooOld = false;
                            }
                        }

                        // Check OS product type for server SKU
                        string productTypeStr = osInfo["ProductType"]?.ToString() ?? "";
                        if (uint.TryParse(productTypeStr, out uint productType))
                        {
                            // ProductType 1 indicates a client OS, while 2 and 3 indicate server OS
                            if (productType != 1)
                            {
                                isOSNotServer = false;
                            }
                        }

                        // Dispose of the ManagementObject to free resources
                        osInfo.Dispose();

                        break;
                    }

                    // Connect to the virtualization namespace to check Hyper-V presence
                    machine.Connect("root\\virtualization\\v2");
                    ManagementObjectCollection hyerpvObjects = machine.GetObjects("Msvm_ComputerSystem", "*");
                    if (hyerpvObjects != null)
                    {
                        // If we can retrieve Hyper-V objects, Hyper-V is enabled
                        isHyperVDisabled = false;
                    }
                });
            }
            catch (Exception ex)
            {
                // In case any problem occur while connecting,
                (new ExceptionView()).HandleException(ex);
                StatusBarChangeBehaviour(false, "Error");
            }
            finally
            {
                // Show warnings based on compatibility checks
                if (isHyperVDisabled)
                {
                    _ = MessageBox.Show(
                            $"Hyper-V is not present on {machine.GetComputerName()}. You can still control the Hyper-V server remotely.",
                            "Warning",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning
                        );
                }

                // Show warning if OS is too old
                if (isOSTooOld)
                {
                    _ = MessageBox.Show(
                                    "Your Windows host is too old to use Discrete Device Assignment. Please consider to upgrade!",
                                    "Warning",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Warning
                                );
                }

                // Show warning if OS is not server SKU
                if (isOSNotServer)
                {
                    _ = MessageBox.Show(
                                "Your SKU of Windows may not support Discrete Device Assignment. Please use the \"Server\" SKU.",
                                "Warning",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning
                            );
                }
            }

            // Refresh the VM list after reinitializing the connection
            if (!isHyperVDisabled)
            {
                await RefreshVMs();
            }
        }

        /*
         * Refresh the VM list from the Hyper-V host
         */
        private async Task RefreshVMs()
        {
            // Clear existing VM and device lists
            VMList.Items.Clear();
            DevicePerVMList.Items.Clear();

            StatusBarChangeBehaviour(true, "Refreshing");

            // Try to retrieve VM information and update the lists
            try
            {
                await Task.Run(() =>
                {
                    // Clear existing VM objects
                    vmObjects.Clear();

                    // Connect to the Hyper-V WMI namespace
                    machine.Connect("root\\virtualization\\v2");
                    foreach (ManagementObject vmInfo in machine.GetObjects("Msvm_ComputerSystem", "Caption, ElementName, EnabledState, Name").Cast<ManagementObject>())
                    {
                        // Retrieve VM properties
                        string vmInfoCaption = vmInfo["Caption"]?.ToString();
                        string vmInfoName = vmInfo["Name"]?.ToString();
                        string vmInfoElementName = vmInfo["ElementName"]?.ToString();
                        ushort vmInfoState = (ushort)vmInfo["EnabledState"];

                        // Validate retrieved properties
                        if (vmInfoCaption == null || vmInfoName == null || vmInfoElementName == null || vmInfoState > 10)
                        {
                            vmInfo.Dispose();
                            continue;
                        }

                        // Map VM state to human-readable status
                        if (vmInfoCaption.Equals("Virtual Machine"))
                        {
                            vmObjects[vmInfoName] = (vmInfoElementName, vmStatusMap[vmInfoState], new List<(string instanceId, string devInstancePath)>());
                        }

                        // Dispose of the ManagementObject to free resources
                        vmInfo.Dispose();
                    }

                    // Retrieve PCI Express Setting Data to find assigned devices
                    foreach (ManagementObject deviceid in machine.GetObjects("Msvm_PciExpressSettingData", "Caption, HostResource, InstanceID").Cast<ManagementObject>())
                    {
                        // Retrieve the Caption property
                        string pciSettingCaption = deviceid["Caption"]?.ToString();

                        if (pciSettingCaption == null)
                        {
                            deviceid.Dispose();
                            continue;
                        }

                        //  If the device is a PCI Express Port, process it
                        if (pciSettingCaption.Equals("PCI Express Port"))
                        {
                            string instanceId = deviceid["InstanceID"]?.ToString();

                            if (instanceId == null)
                            {
                                deviceid.Dispose();
                                continue;
                            }

                            string[] payload = instanceId.Replace("Microsoft:", "").Split('\\');
                            string[] hostResources = (string[])deviceid["HostResource"];

                            if (hostResources == null)
                            {
                                deviceid.Dispose();
                                continue;
                            }

                            if (vmObjects.ContainsKey(payload[0]))
                            {
                                string devPath = Regex.Replace(hostResources[0].Split(',')[1], "DeviceID=\"Microsoft:[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-4[0-9a-fA-F]{3}-[89aAbB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}\\\\\\\\", "").Replace("\"", "").Replace("\\\\", "\\");
                                vmObjects[payload[0]].devices.Add((instanceId, devPath));
                            }

                            deviceid.Dispose();
                        }
                    }

                    // Populate the VMList UI element with the retrieved VM information
                    foreach (KeyValuePair<string, (string vmName, string vmStatus, List<(string instanceId, string devInstancePath)> devices)> vmObject in vmObjects)
                    {
                        // Use Dispatcher to update the UI thread
                        Dispatcher.Invoke(() =>
                        {
                            _ = VMList.Items.Add(new
                            {
                                VMId = vmObject.Key,
                                VMName = vmObject.Value.vmName,
                                VMStatus = vmObject.Value.vmStatus
                            });
                        });
                    }
                });
                
                StatusBarChangeBehaviour(false, "Done");
            }
            catch (Exception ex)
            {
                // In case any problem occur while refreshing,
                (new ExceptionView()).HandleException(ex);
                StatusBarChangeBehaviour(false, "Error");
            }
        }

        /*
         * Change the status bar behaviour
         */
        private void StatusBarChangeBehaviour(bool isRefresh, string labelMessage = "Refreshing")
        {
            BottomProgressBarStatus.Visibility = (isRefresh) ? Visibility.Visible : Visibility.Hidden;
            BottomProgressBarStatus.IsIndeterminate = isRefresh;
            BottomLabelStatus.Text = labelMessage;
        }
    }
}