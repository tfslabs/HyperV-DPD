using DDAGUI.WMIProperties;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Security;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace DDAGUI
{
    public partial class MainWindow : Window
    {
        /*
         *  Global properties
         */
        private MachineMethods machine;
        private readonly Dictionary<string, (string vmName, string vmStatus, List<(string instanceId, string devInstancePath)> devices)> vmObjects;

        public MainWindow()
        {
            InitializeComponent();

            machine = new MachineMethods();
            vmObjects = new Dictionary<string, (string vmName, string vmStatus, List<(string instanceId, string devInstancePath)> devices)>();

            Loaded += MainWindow_Loaded;
        }

        public async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await ReinitializeConnection();
        }

        /*
         * Button and UI methods behaviour....
         */

        private async void ConnectToAnotherComputer_Button(object sender, RoutedEventArgs e)
        {
            StatusBarChangeBehaviour(true, "Connecting");

            ConnectForm connectForm = new ConnectForm();

            (string computerName, string username, SecureString password) userCredential = connectForm.ReturnValue();

            if (userCredential.computerName != "")
            {
                machine = (userCredential.computerName.Equals("localhost")) ?
                    new MachineMethods() :
                    new MachineMethods(userCredential);
                await ReinitializeConnection();
            }
        }

        private void QuitMainWindow_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private async void AddDevice_Click(object sender, RoutedEventArgs e)
        {
            if (VMList.SelectedItem != null)
            {
                StatusBarChangeBehaviour(true, "Adding device");

                AddDevice addDevice = new AddDevice(machine);
                string deviceId = addDevice.GetDeviceId();

                string vmName = VMList.SelectedItem.GetType().GetProperty("VMName").GetValue(VMList.SelectedItem, null).ToString();
                try
                {
                    if (deviceId != null)
                    {
                        await Task.Run(() =>
                        {
                            machine.ChangePnpDeviceBehaviour(deviceId, "Disable");
                            machine.MountPnPDeviceToPcip(deviceId);
                            machine.MountIntoVM(vmName, deviceId);
                        });

                        await RefreshVMs();
                    }
                    else
                    {
                        StatusBarChangeBehaviour(false, "No Device Added");
                    }
                }
                catch (Exception ex)
                {
                    WMIDefaultValues.HandleException(ex, machine.GetComputerName());
                    StatusBarChangeBehaviour(false, "Error, re-enabling device");

                    // Re-enable target device
                    try
                    {
                        await Task.Run(() =>
                        {
                            machine.ChangePnpDeviceBehaviour(deviceId, "Enable");
                        });
                    }
                    catch (Exception exp)
                    {
                        WMIDefaultValues.HandleException(exp, machine.GetComputerName());
                        StatusBarChangeBehaviour(false, "Re-enable device failed");
                    }
                }
            }
            else
            {
                MessageBox.Show(
                    "Please select a virtual machine on the list",
                    "Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                StatusBarChangeBehaviour(false, "No Device Added");
            }
        }

        private async void RemDevice_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (VMList.SelectedItem != null && DevicePerVMList.SelectedItem != null)
                {
                    StatusBarChangeBehaviour(true, "Removing device");
                    string deviceId = DevicePerVMList.SelectedItem.GetType().GetProperty("DeviceID").GetValue(DevicePerVMList.SelectedItem, null).ToString();
                    string devicePath = DevicePerVMList.SelectedItem.GetType().GetProperty("DevicePath").GetValue(DevicePerVMList.SelectedItem, null).ToString();

                    await Task.Run(() =>
                    {
                        machine.DismountFromVM(deviceId);
                        machine.DismountPnPDeviceFromPcip(devicePath);
                        machine.ChangePnpDeviceBehaviour(deviceId.Replace("PCIP\\", "PCI\\"), "Enable");
                    });

                    await RefreshVMs();
                }
                else
                {
                    MessageBox.Show(
                        "Please select a device from the virtual machine List!",
                        "Warning",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning
                    );
                    StatusBarChangeBehaviour(false, "No Device Selected");
                }
            }
            catch (Exception ex)
            {
                WMIDefaultValues.HandleException(ex, machine.GetComputerName());
                StatusBarChangeBehaviour(false, "Error");
            }
        }

        private void CopyDevAddress_Click(object sender, RoutedEventArgs e)
        {
            if (VMList.SelectedItems != null && DevicePerVMList.SelectedItems != null)
            {
                try
                {
                    if (VMList.SelectedItem != null && DevicePerVMList.SelectedItem != null)
                    {
                        string devicePath = DevicePerVMList.SelectedItem.GetType().GetProperty("DevicePath").GetValue(DevicePerVMList.SelectedItem, null).ToString();
                        Clipboard.SetText(devicePath);
                        MessageBox.Show(
                            $"Copied device Path {devicePath} into clipboard",
                            "Info",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information
                        );
                    }
                    else
                    {
                        MessageBox.Show(
                            "Please select a device from the virtual machine List!",
                            "Warning",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning
                        );
                    }
                }
                catch (Exception ex)
                {
                    WMIDefaultValues.HandleException(ex, machine.GetComputerName());
                    StatusBarChangeBehaviour(false, "Error");
                }
            }
            else
            {
                MessageBox.Show(
                    "Please select a virtual machine and a device!",
                    "Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                StatusBarChangeBehaviour(false, "No Device Selected");
            }
        }

        private async void ChangeMemLocation_Click(object sender, RoutedEventArgs e)
        {
            if (VMList.SelectedItem != null)
            {
                try
                {
                    string vmName = VMList.SelectedItem.GetType().GetProperty("VMName").GetValue(VMList.SelectedItem, null).ToString();

                    ChangeMemorySpace changeMemorySpace = new ChangeMemorySpace(vmName);
                    (UInt64 lowMem, UInt64 highMem) = changeMemorySpace.ReturnValue();

                    if (lowMem != 0 && highMem != 0)
                    {
                        await Task.Run(() =>
                        {
                            machine.ChangeMemAllocate(vmName, lowMem, highMem);
                        });

                        await RefreshVMs();
                    }
                }
                catch (Exception ex)
                {
                    WMIDefaultValues.HandleException(ex, machine.GetComputerName());
                    StatusBarChangeBehaviour(false, "Error");
                }
            }
            else
            {
                MessageBox.Show(
                    "Please select a virtual machine on the list",
                    "Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                StatusBarChangeBehaviour(false, "No virtual machine Selected");
            }
        }

        private void HyperVServiceStatus_Click(object sender, RoutedEventArgs e)
        {
            HyperVStatus hyperVStatus = new HyperVStatus(machine);
            hyperVStatus.ShowDialog();
        }

        private void AssignableDevice_Click(object sender, RoutedEventArgs e)
        {
            CheckForAssignableDevice device = new CheckForAssignableDevice(machine);
            device.ShowDialog();
        }

        private async void RemAllDevice_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MessageBoxResult isRemoveAll = MessageBox.Show(
                            $"Do you want to remove all assigned devices? (This option should only be used in case you accidently remove a virtual machine without unmounting it first)",
                            "Remove all devices",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question
                );

                if (isRemoveAll == MessageBoxResult.Yes)
                {
                    StatusBarChangeBehaviour(true, "In action");

                    await Task.Run(() =>
                    {
                        machine.Connect("root\\virtualization\\v2");
                        ManagementObjectCollection devSettings = machine.GetObjects("Msvm_PciExpressSettingData", "InstanceID");
                        ManagementObjectCollection devMount = machine.GetObjects("Msvm_PciExpress", "DeviceInstancePath");

                        if (!(devSettings == null || devSettings.Count == 0))
                        {
                            foreach (ManagementObject deviceid in devSettings.Cast<ManagementObject>())
                            {
                                string devInstanceId = deviceid["InstanceID"]?.ToString();

                                if (devInstanceId == null || devInstanceId.Length == 0)
                                {
                                    continue;
                                }

                                if (!devInstanceId.Contains("Microsoft:Definition"))
                                {
                                    machine.DismountFromVM(devInstanceId);
                                }
                            }
                        }

                        if (!(devMount == null || devMount.Count == 0))
                        {
                            foreach (ManagementObject devInstancePath in devMount.Cast<ManagementObject>())
                            {
                                string devInsPath = devInstancePath["DeviceInstancePath"].ToString();
                                machine.DismountPnPDeviceFromPcip(devInsPath);
                                machine.ChangePnpDeviceBehaviour(devInsPath.Replace("PCIP\\", "PCI\\"), "Enable");
                            }
                        }
                    });

                    await RefreshVMs();
                }
                else
                {
                    StatusBarChangeBehaviour(false, "No Action Performed");
                }
            }
            catch (Exception ex)
            {
                WMIDefaultValues.HandleException(ex, machine.GetComputerName());
                StatusBarChangeBehaviour(false, "Error");
            }
        }

        private async void HyperVRefresh_Click(object sender, RoutedEventArgs e)
        {
            await RefreshVMs();
        }

        private void AboutBox_Click(object sender, RoutedEventArgs e)
        {
            About aboutbox = new About();
            aboutbox.ShowDialog();
        }

        private async void ChangeCacheTypes_Click(object sender, RoutedEventArgs e)
        {
            if (VMList.SelectedItem != null)
            {
                try
                {
                    string vmName = VMList.SelectedItem.GetType().GetProperty("VMName").GetValue(VMList.SelectedItem, null).ToString();
                    MessageBoxResult isEnableMemCache = MessageBox.Show(
                            $"Do you want to enable guest control cache type for {vmName}? (Yes to enable, No to disable, Cancel to leave it as is)",
                            "Enable Guest Control Cache",
                            MessageBoxButton.YesNoCancel,
                            MessageBoxImage.Question
                    );

                    if (isEnableMemCache == MessageBoxResult.Yes)
                    {
                        await Task.Run(() =>
                        {
                            machine.ChangeGuestCacheType(vmName, true);
                            MessageBox.Show(
                                $"Enabled guest control cache type successfully for {vmName}",
                                "Info",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information
                            );
                        });
                    }
                    else if (isEnableMemCache == MessageBoxResult.No)
                    {
                        await Task.Run(() =>
                        {
                            machine.ChangeGuestCacheType(vmName, false);
                            MessageBox.Show(
                                $"Disabled guest control cache type successfully for {vmName}",
                                "Info",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information
                            );
                        });
                    }

                    await RefreshVMs();
                }
                catch (Exception ex)
                {
                    WMIDefaultValues.HandleException(ex, machine.GetComputerName());
                    StatusBarChangeBehaviour(false, "Error");
                }
            }
            else
            {
                MessageBox.Show(
                    "Please select a virtual machine to add a device",
                    "Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
            }
        }

        private async void VMList_Select(object sender, RoutedEventArgs e)
        {
            if (VMList.SelectedItem != null)
            {
                DevicePerVMList.Items.Clear();
                string vmId = VMList.SelectedItem.GetType().GetProperty("VMId").GetValue(VMList.SelectedItem, null).ToString();
                await Task.Run(() =>
                {
                    foreach (KeyValuePair<string, (string vmName, string vmStatus, List<(string instanceId, string devInstancePath)> devices)> vmObject in vmObjects)
                    {
                        if (vmObject.Key.Equals(vmId))
                        {
                            foreach ((string instanceId, string devInstancePath) in vmObject.Value.devices)
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    DevicePerVMList.Items.Add(new
                                    {
                                        DeviceID = instanceId,
                                        DevicePath = devInstancePath
                                    });
                                });
                            }
                            break;
                        }
                    }
                });
            }
        }

        /*
         * Non-button methods
         */

        private async Task ReinitializeConnection()
        {
            StatusBarChangeBehaviour(true, "Connecting");

            try
            {

                await Task.Run(() =>
                {
                    machine.Connect("root\\virtualization\\v2");
                    ManagementObjectCollection hyerpvObjects = machine.GetObjects("Msvm_ComputerSystem", "*");
                    if (hyerpvObjects == null || hyerpvObjects.Count == 0)
                    {
                        MessageBox.Show(
                            $"Hyper-V is not present on {machine.GetComputerName()}. You can still control the Hyper-V server remotely.",
                            "Warning",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning
                        );
                    }

                    machine.Connect("root\\cimv2");
                    foreach (ManagementObject osInfo in machine.GetObjects("Win32_OperatingSystem", "BuildNumber, ProductType").Cast<ManagementObject>())
                    {
                        int buildNumber = int.Parse(osInfo["BuildNumber"]?.ToString());
                        if (buildNumber < 14393)
                        {
                            MessageBox.Show(
                                "Your Windows host is too old to use Discrete Device Assignment. Please consider to upgrade!",
                                "Warning",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning
                            );
                        }

                        UInt32 productType = (UInt32)osInfo["ProductType"];
                        if (productType == (UInt32)1)
                        {
                            MessageBox.Show(
                                "Your SKU of Windows may not support Discrete Device Assignment. Please use the \"Server\" SKU.",
                                "Warning",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning
                            );
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                WMIDefaultValues.HandleException(ex, machine.GetComputerName());
                StatusBarChangeBehaviour(false, "Error");
            }

            await RefreshVMs();
        }

        private async Task RefreshVMs()
        {
            // Clear variables, then the GUI
            vmObjects.Clear();
            VMList.Items.Clear();
            DevicePerVMList.Items.Clear();

            StatusBarChangeBehaviour(true, "Refreshing");

            try
            {
                await Task.Run(() =>
                {
                    machine.Connect("root\\virtualization\\v2");

                    foreach (ManagementObject vmInfo in machine.GetObjects("Msvm_ComputerSystem", "Caption, ElementName, EnabledState, Name").Cast<ManagementObject>())
                    {
                        if (vmInfo["Caption"].ToString().Equals("Virtual Machine"))
                        {
                            vmObjects[vmInfo["Name"].ToString()] = (vmInfo["ElementName"].ToString(), WMIDefaultValues.vmStatusMap[int.Parse(vmInfo["EnabledState"]?.ToString() ?? "0")], new List<(string instanceId, string devInstancePath)>());
                        }
                    }

                    foreach (ManagementObject deviceid in machine.GetObjects("Msvm_PciExpressSettingData", "Caption, HostResource, InstanceID").Cast<ManagementObject>())
                    {
                        if (deviceid["Caption"].ToString().Equals("PCI Express Port"))
                        {
                            string instanceId = deviceid["InstanceID"].ToString();

                            string hostResources = ((string[])deviceid["HostResource"])[0];
                            string[] payload = instanceId.Replace("Microsoft:", "").Split('\\');

                            if (vmObjects.ContainsKey(payload[0]))
                            {
                                string devPath = Regex.Replace(((string[])hostResources.Split(','))[1], "DeviceID=\"Microsoft:[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-4[0-9a-fA-F]{3}-[89aAbB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}\\\\\\\\", "").Replace("\"", "").Replace("\\\\", "\\");
                                vmObjects[payload[0]].devices.Add((instanceId, devPath));
                            }
                        }
                    }

                    foreach (var vmObject in vmObjects)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            VMList.Items.Add(new
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
                WMIDefaultValues.HandleException(ex, machine.GetComputerName());
                StatusBarChangeBehaviour(false, "Error");
            }
        }

        private void StatusBarChangeBehaviour(bool isRefresh, string labelMessage = "Refreshing")
        {
            BottomProgressBarStatus.Visibility = (isRefresh) ? Visibility.Visible : Visibility.Hidden;
            BottomProgressBarStatus.IsIndeterminate = isRefresh;
            BottomLabelStatus.Text = labelMessage;
        }
    }
}