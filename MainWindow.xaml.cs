using DDAGUI.WMIProperties;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Security;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace DDAGUI
{
    public partial class MainWindow : Window
    {
        /*
         * Msvm_PciExpress for Get-VMHostAssignableDevice
         * Msvm_PciExpressSettingData for Get-VMAssignableDevice -VMName $vm
         */

        /*
         *  Global properties
         */
        private MachineMethods machine;
        private Dictionary<string, (string vmName, string vmStatus, List<string> devices)> vmObjects;

        public MainWindow()
        {
            InitializeComponent();

            StatusBarChangeBehaviour(true);

            machine = new MachineMethods();
            vmObjects = new Dictionary<string, (string vmName, string vmStatus, List<string> devices)>();

            Loaded += MainWindow_Loaded;

            StatusBarChangeBehaviour(false, "Done");
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
                machine = (userCredential.computerName.Equals("localhost")) ? new MachineMethods() : new MachineMethods(userCredential);
                await ReinitializeConnection();
            }

            StatusBarChangeBehaviour(false, "Done");
        }

        private void QuitMainWindow_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void AddDevice_Click(object sender, RoutedEventArgs e)
        {
            if (VMList.SelectedItem != null)
            {
                try
                {
                    StatusBarChangeBehaviour(true, "Adding devices");

                    AddDevice addDevice = new AddDevice(machine);

                    string deviceId = addDevice.GetDeviceId();
                    string vmName = VMList.SelectedItem.GetType().GetProperty("VMName").GetValue(VMList.SelectedItem, null).ToString();

                    MessageBox.Show($"Add device {deviceId} into {vmName}", "Info", MessageBoxButton.OK, MessageBoxImage.Information);

                    StatusBarChangeBehaviour(false, "Done");
                }
#if DEBUG
                catch (NullReferenceException ex)
                {

                    MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusBarChangeBehaviour(false, "No Device Added");
#else
                catch (NullReferenceException)
                {
                    StatusBarChangeBehaviour(false, "No Device Added");
#endif
                }
            }
            else
            {
                MessageBox.Show("Please select a VM to add a device!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void RemDevice_Click(object sender, RoutedEventArgs e)
        {
            if (VMList.SelectedItems != null && DevicePerVMList.SelectedItems != null)
            {
                try
                {
                    string vmName = VMList.SelectedItem.GetType().GetProperty("VMName").GetValue(VMList.SelectedItem, null).ToString();
                    string deviceId = DevicePerVMList.SelectedItem.GetType().GetProperty("DeviceID").GetValue(DevicePerVMList.SelectedItem, null).ToString();

                    MessageBox.Show($"Remove device {deviceId} from {vmName}", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (UnauthorizedAccessException ex)
                {
#if DEBUG
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#else
                    MessageBox.Show($"Failed to catch the Authenticate with {machine.GetComputerName()}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
                }
                catch (COMException ex)
                {
#if DEBUG
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#else
                    MessageBox.Show($"Failed to reach {machine.GetComputerName()}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
                }
                catch (ManagementException ex)
                {
#if DEBUG
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#else
                    MessageBox.Show($"Failed to catch the Management Method: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
                }
#if DEBUG
                catch (NullReferenceException ex)
                {

                    MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusBarChangeBehaviour(false, "No Device Selected");
#else
                catch (NullReferenceException)
                {
                    StatusBarChangeBehaviour(false, "No Device Selected");
#endif
                }
            }
            else
            {
                MessageBox.Show("Please select a VM and a device!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void CopyDevAddress_Click(object sender, RoutedEventArgs e)
        {
            if (VMList.SelectedItems != null && DevicePerVMList.SelectedItems != null)
            {
                try
                {
                    string deviceId = DevicePerVMList.SelectedItem.GetType().GetProperty("DeviceID").GetValue(DevicePerVMList.SelectedItem, null).ToString();
                    Clipboard.SetText(deviceId);
                    MessageBox.Show($"Copied device ID {deviceId} into clipboard", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                }
#if DEBUG
                catch (NullReferenceException ex)
                {

                    MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusBarChangeBehaviour(false, "No Device Selected");
#else
                catch (NullReferenceException)
                {
                    StatusBarChangeBehaviour(false, "No Device Selected");
#endif
                }
            }
            else
            {
                MessageBox.Show("Please select a VM and a device!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ChangeMemLocation_Click(object sender, RoutedEventArgs e)
        {

        }

        private void HyperVServiceStatus_Click(object sender, RoutedEventArgs e)
        {
            HyperVStatus hyperVStatus = new HyperVStatus(machine);
            hyperVStatus.ShowDialog();
        }

        private void AssignableDevice_Click(object sender, RoutedEventArgs e)
        {

        }

        private void RemAllDevice_Click(object sender, RoutedEventArgs e)
        {

        }

        private async void HyperVRefresh_Click(object sender, RoutedEventArgs e)
        {
            StatusBarChangeBehaviour(true);

            await RefreshVMs();

            StatusBarChangeBehaviour(false, "Done");
        }

        private void AboutBox_Click(object sender, RoutedEventArgs e)
        {
            About aboutbox = new About();
            aboutbox.ShowDialog();
        }

        private void ChangeCacheTypes_Click(object sender, RoutedEventArgs e)
        {
            if (VMList.SelectedItem != null)
            {
                try
                {
                    string vmName = VMList.SelectedItem.GetType().GetProperty("VMName").GetValue(VMList.SelectedItem, null).ToString();
                    MessageBoxResult isEnableMemCache = MessageBox.Show(
                            $"Do you want to enable guest control cache type for {vmName}? (Yes to enable, No to disable, Cancel to leave it as is)",
                            "Enable Guest Control Cache", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

                    if (isEnableMemCache == MessageBoxResult.Yes)
                    {

                    }
                    else if (isEnableMemCache == MessageBoxResult.No)
                    {

                    }
                    else
                    {

                    }
                }
                catch (UnauthorizedAccessException ex)
                {
#if DEBUG
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#else
                    MessageBox.Show($"Failed to catch the Authenticate with {machine.GetComputerName()}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
                }
                catch (COMException ex)
                {
#if DEBUG
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#else
                    MessageBox.Show($"Failed to reach {machine.GetComputerName()}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
                }
                catch (ManagementException ex)
                {
#if DEBUG
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#else
                    MessageBox.Show($"Failed to catch the Management Method: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
                }
            }
            else
            {
                MessageBox.Show("Please select a VM to add a device", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                    foreach (var vmObject in vmObjects)
                    {
                        if (vmObject.Key.Equals(vmId))
                        {
                            foreach (string deviceInstanceid in vmObject.Value.devices)
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    DevicePerVMList.Items.Add(new
                                    {
                                        DeviceID = deviceInstanceid
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
            try
            {
                machine.Connect("root\\virtualization\\v2");
                var hyerpvObjects = machine.GetObjects("Msvm_ComputerSystem", "*");
                if (hyerpvObjects == null || hyerpvObjects.Count == 0)
                {
                    MessageBox.Show("Please ensure Hyper-V is installed and running on this machine. You can still control the Hyper-V server remotely.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                machine.Connect("root\\cimv2");
                var getMachineInfo = machine.GetObjects("Win32_OperatingSystem", "BuildNumber, Caption");
                foreach (var osInfo in getMachineInfo)
                {
                    string osName = osInfo["Caption"]?.ToString();
                    int buildNumber = int.Parse(osInfo["BuildNumber"]?.ToString());

                    if (buildNumber < 14393)
                    {
                        MessageBox.Show("Your Windows host is too old to use Hyper-V DDA. Please consider to upgrade!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }

                    if (!osName.Trim().ToLower().Contains("server"))
                    {
                        MessageBox.Show("Your SKU of Windows may not support Hyper-V with DDA. Please use the SKU \"Server\".", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (UnauthorizedAccessException ex)
            {
#if DEBUG
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#else
                MessageBox.Show($"Failed to catch the Authenticate with {machine.GetComputerName()}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
            }
            catch (COMException ex)
            {
#if DEBUG
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#else
                MessageBox.Show($"Failed to reach {machine.GetComputerName()}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
            }
            catch (ManagementException ex)
            {
#if DEBUG
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#else
                MessageBox.Show($"Failed to catch the Management Method: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
            }

            await RefreshVMs();
        }

        private async Task RefreshVMs()
        {
            VMList.Items.Clear();
            DevicePerVMList.Items.Clear();

            try
            {
                await Task.Run(() =>
                {
                    machine.Connect("root\\virtualization\\v2");

                    foreach (ManagementObject vmInfo in machine.GetObjects("Msvm_ComputerSystem", "Caption, ElementName, EnabledState, Name").Cast<ManagementObject>())
                    {
                        if (vmInfo["Caption"].ToString().Equals("Virtual Machine"))
                        {
                            string vmName = vmInfo["ElementName"].ToString();
                            string vmStatus = WMIDefaultValues.vmStatusMap[int.Parse(vmInfo["EnabledState"]?.ToString() ?? "0")];
                            vmObjects[vmInfo["Name"].ToString()] = (vmName, vmStatus, new List<string>());
                        }
                    }

                    foreach (ManagementObject deviceid in machine.GetObjects("Msvm_PciExpressSettingData", "Caption, HostResource, InstanceID").Cast<ManagementObject>())
                    {
                        if (deviceid["Caption"].ToString().Equals("PCI Express Port"))
                        {
                            string hostResources = ((string[])deviceid["HostResource"])[0];
                            string[] payload = deviceid["InstanceID"].ToString().Replace("Microsoft:", "").Split('\\');

                            if (vmObjects.ContainsKey(payload[0]))
                            {
                                string key = Regex.Replace(((string[])hostResources.Split(','))[1], "DeviceID=\"Microsoft:[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-4[0-9a-fA-F]{3}-[89aAbB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}\\\\\\\\", "").Replace("\"", "").Replace("\\\\", "\\");
                                vmObjects[payload[0]].devices.Add(key);
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
            }
            catch (UnauthorizedAccessException ex)
            {
#if DEBUG
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#else
                MessageBox.Show($"Failed to catch the Authenticate with {machine.GetComputerName()}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
            }
            catch (COMException ex)
            {
#if DEBUG
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#else
                MessageBox.Show($"Failed to reach {machine.GetComputerName()}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
            }
            catch (ManagementException ex)
            {
#if DEBUG
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#else
                MessageBox.Show($"Failed to catch the Management Method: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
            }
        }

        private void StatusBarChangeBehaviour(bool isRefresh, string labelMessage = "Refreshing...")
        {
            BottomProgressBarStatus.IsIndeterminate = isRefresh;
            BottomLabelStatus.Text = labelMessage;
        }
    }
}
