using TheFlightSims.HyperVDPD.WMIProperties;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Security;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace TheFlightSims.HyperVDPD
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
                machine = (string.Equals(userCredential.computerName, "localhost", StringComparison.OrdinalIgnoreCase)) ? new MachineMethods() : new MachineMethods(userCredential);
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
                string vmName = VMList.SelectedItem.GetType().GetProperty("VMName").GetValue(VMList.SelectedItem, null)?.ToString();

                if (deviceId != null && vmName != null)
                {
                    try
                    {
                        await Task.Run(() =>
                        {
                            machine.ChangePnpDeviceBehaviour(deviceId, "Disable");
                            machine.MountPnPDeviceToPcip(deviceId);
                            machine.MountIntoVM(vmName, deviceId);
                        });

                        await RefreshVMs();
                    }
                    catch (Exception ex)
                    {
                        WMIDefaultValues.HandleException(ex, machine.GetComputerName());
                        StatusBarChangeBehaviour(false, "Error, Re-enable device");

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
                    StatusBarChangeBehaviour(false, "No Device Added");
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

        private async void RemDevice_Click(object sender, RoutedEventArgs e)
        {
            if (VMList.SelectedItem != null && DevicePerVMList.SelectedItem != null)
            {
                try
                {
                    StatusBarChangeBehaviour(true, "Removing device");
                    string deviceId = DevicePerVMList.SelectedItem.GetType().GetProperty("DeviceID").GetValue(DevicePerVMList.SelectedItem, null)?.ToString();
                    string devicePath = DevicePerVMList.SelectedItem.GetType().GetProperty("DevicePath").GetValue(DevicePerVMList.SelectedItem, null)?.ToString();

                    if (deviceId != null && devicePath != null)
                    {
                        await Task.Run(() =>
                        {
                            machine.DismountFromVM(deviceId);
                            machine.DismountPnPDeviceFromPcip(devicePath);
                            machine.ChangePnpDeviceBehaviour(devicePath.Replace("PCIP\\", "PCI\\"), "Enable");
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
                    "Please select a device from the virtual machine list!",
                    "Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );

                StatusBarChangeBehaviour(false, "No Device Selected");
            }
        }

        private void CopyDevAddress_Click(object sender, RoutedEventArgs e)
        {
            if (VMList.SelectedItem != null && DevicePerVMList.SelectedItem != null)
            {
                string devicePath = DevicePerVMList.SelectedItem.GetType().GetProperty("DevicePath").GetValue(DevicePerVMList.SelectedItem, null)?.ToString();

                if (devicePath != null)
                {
                    Clipboard.SetText(devicePath);
                    MessageBox.Show(
                        $"Copied device Path {devicePath} into clipboard",
                        "Info",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information
                    );
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

                        using (ManagementObjectCollection devSettings = machine.GetObjects("Msvm_PciExpressSettingData", "InstanceID"))
                        {
                            if (devSettings != null)
                            {
                                foreach (ManagementObject deviceid in devSettings.Cast<ManagementObject>())
                                {
                                    string devInstanceId = deviceid["InstanceID"]?.ToString();

                                    if (devInstanceId == null)
                                    {
                                        deviceid.Dispose();
                                        continue;
                                    }

                                    if (!devInstanceId.Contains("Microsoft:Definition"))
                                    {
                                        machine.DismountFromVM(devInstanceId);
                                    }

                                    deviceid.Dispose();
                                }
                            }
                        }

                        using (ManagementObjectCollection devMount = machine.GetObjects("Msvm_PciExpress", "DeviceInstancePath"))
                        {
                            if (devMount != null)
                            {
                                foreach (ManagementObject devInstancePath in devMount.Cast<ManagementObject>())
                                {
                                    string devInsPath = devInstancePath["DeviceInstancePath"]?.ToString();

                                    if (devInsPath != null)
                                    {
                                        machine.DismountPnPDeviceFromPcip(devInsPath);
                                        machine.ChangePnpDeviceBehaviour(devInsPath.Replace("PCIP\\", "PCI\\"), "Enable");
                                        devInstancePath.Dispose();
                                    }
                                    else
                                    {
                                        devInstancePath.Dispose();
                                        continue;
                                    }
                                }
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
                    string vmName = VMList.SelectedItem.GetType().GetProperty("VMName").GetValue(VMList.SelectedItem, null)?.ToString();
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
                        });

                        MessageBox.Show(
                            $"Enabled guest control cache type successfully for {vmName}",
                            "Info",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information
                        );
                    }
                    else if (isEnableMemCache == MessageBoxResult.No)
                    {
                        await Task.Run(() =>
                        {
                            machine.ChangeGuestCacheType(vmName, false);
                        });

                        MessageBox.Show(
                                $"Disabled guest control cache type successfully for {vmName}",
                                "Info",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information
                        );
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
                string vmId = VMList.SelectedItem.GetType().GetProperty("VMId").GetValue(VMList.SelectedItem, null)?.ToString();
                if (vmId != null)
                {
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
        }

        /*
         * Non-button methods
         */

        private async Task ReinitializeConnection()
        {
            bool isHyperVDisabled = true, isOSTooOld = true, isOSNotServer = true;

            StatusBarChangeBehaviour(true, "Connecting");

            try
            {
                await Task.Run(() =>
                {
                    machine.Connect("root\\cimv2");
                    foreach (ManagementObject osInfo in machine.GetObjects("Win32_OperatingSystem", "BuildNumber, ProductType").Cast<ManagementObject>())
                    {
                        string osInfoBuildNumber = osInfo["BuildNumber"]?.ToString() ?? "";
                        if (int.TryParse(osInfoBuildNumber, out int buildNumber))
                        {
                            if (buildNumber >= 14393)
                            {
                                isOSTooOld = false;
                            }
                        }

                        string productTypeStr = osInfo["ProductType"]?.ToString() ?? "";
                        if (UInt32.TryParse(productTypeStr, out UInt32 productType))
                        {
                            if (productType != (UInt32)1)
                            {
                                isOSNotServer = false;
                            }
                        }

                        osInfo.Dispose();

                        break;
                    }

                    machine.Connect("root\\virtualization\\v2");
                    ManagementObjectCollection hyerpvObjects = machine.GetObjects("Msvm_ComputerSystem", "*");
                    if (hyerpvObjects != null)
                    {
                        isHyperVDisabled = false;
                    }
                });
            }
            catch (Exception ex)
            {
                WMIDefaultValues.HandleException(ex, machine.GetComputerName());
                StatusBarChangeBehaviour(false, "Error");
            }
            finally
            {
                if (isHyperVDisabled)
                {
                    MessageBox.Show(
                            $"Hyper-V is not present on {machine.GetComputerName()}. You can still control the Hyper-V server remotely.",
                            "Warning",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning
                        );
                }

                if (isOSTooOld)
                {
                    MessageBox.Show(
                                    "Your Windows host is too old to use Discrete Device Assignment. Please consider to upgrade!",
                                    "Warning",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Warning
                                );
                }

                if (isOSNotServer)
                {
                    MessageBox.Show(
                                "Your SKU of Windows may not support Discrete Device Assignment. Please use the \"Server\" SKU.",
                                "Warning",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning
                            );
                }
            }

            await RefreshVMs();
        }

        private async Task RefreshVMs()
        {
            VMList.Items.Clear();
            DevicePerVMList.Items.Clear();

            StatusBarChangeBehaviour(true, "Refreshing");

            try
            {
                await Task.Run(() =>
                {
                    vmObjects.Clear();

                    machine.Connect("root\\virtualization\\v2");
                    foreach (ManagementObject vmInfo in machine.GetObjects("Msvm_ComputerSystem", "Caption, ElementName, EnabledState, Name").Cast<ManagementObject>())
                    {
                        string vmInfoCaption = vmInfo["Caption"]?.ToString();
                        string vmInfoName = vmInfo["Name"]?.ToString();
                        string vmInfoElementName = vmInfo["ElementName"]?.ToString();
                        UInt16 vmInfoState = (UInt16)vmInfo["EnabledState"];

                        if (vmInfoCaption == null || vmInfoName == null || vmInfoElementName == null || vmInfoState > 10)
                        {
                            vmInfo.Dispose();
                            continue;
                        }

                        if (vmInfoCaption.Equals("Virtual Machine"))
                        {
                            vmObjects[vmInfoName] = (vmInfoElementName, WMIDefaultValues.vmStatusMap[vmInfoState], new List<(string instanceId, string devInstancePath)>());
                        }

                        vmInfo.Dispose();
                    }

                    foreach (ManagementObject deviceid in machine.GetObjects("Msvm_PciExpressSettingData", "Caption, HostResource, InstanceID").Cast<ManagementObject>())
                    {
                        string pciSettingCaption = deviceid["Caption"]?.ToString();

                        if (pciSettingCaption == null)
                        {
                            deviceid.Dispose();
                            continue;
                        }

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
                                string devPath = Regex.Replace(((string[])hostResources[0].Split(','))[1], "DeviceID=\"Microsoft:[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-4[0-9a-fA-F]{3}-[89aAbB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}\\\\\\\\", "").Replace("\"", "").Replace("\\\\", "\\");
                                vmObjects[payload[0]].devices.Add((instanceId, devPath));
                            }

                            deviceid.Dispose();
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