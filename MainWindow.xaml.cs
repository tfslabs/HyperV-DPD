using System;
using System.Collections.Generic;
using System.Management;
using System.Windows;

namespace DDAGUI
{
    public partial class MainWindow : Window
    {
        protected WMIWrapper wmi;
        protected string computerName = "localhost";
        private Dictionary<int, string> vmStatusMap = new Dictionary<int, string>
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

        public MainWindow()
        {
            InitializeComponent();

            StatusBarChangeBehaviour(true);

            this.wmi = new WMIWrapper(computerName);

            try
            {
                var hyerpvObjects = this.wmi.getManagementObjectCollection("Msvm_ComputerSystem", "root\\virtualization\\v2");
                if (hyerpvObjects == null || hyerpvObjects.Count == 0)
                {
                    MessageBox.Show("Please ensure Hyper-V is installed and running on this machine.\nApplication shutting down...", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    Application.Current.Shutdown();
                }

            }
            catch (ManagementException e)
            {
                MessageBox.Show(e.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }

            this.RefreshVMs();

            StatusBarChangeBehaviour(false);
        }

        /*
         * Button and UI methods behaviour....
         */

        private void WMIConnect_Click(object sender, RoutedEventArgs e)
        {
            StatusBarChangeBehaviour(true);
            ConnectForm connectForm = new ConnectForm();
            connectForm.ShowDialog();
            StatusBarChangeBehaviour(false);
        }

        private void QuitMainWindow_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void AddDevice_Click(object sender, RoutedEventArgs e)
        {
            StatusBarChangeBehaviour(true);

            AddDevice addDeviceDialog = new AddDevice(this.wmi);
            string deviceId = addDeviceDialog.GetDeviceId();

            try
            {
                MessageBox.Show($"Device ID: {deviceId}", "Device Added", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (ExceptionHandler) 
            {
            
            }

            StatusBarChangeBehaviour(false);
        }

        private void RemDevice_Click(object sender, RoutedEventArgs e)
        {

        }

        private void CopyDevAddress_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ChangeMemLocation_Click(object sender, RoutedEventArgs e)
        {

        }

        private void IsGuestControlledCacheTypes_Click(object sender, RoutedEventArgs e)
        {

        }

        private void HyperVServiceStatus_Click(object sender, RoutedEventArgs e)
        {

        }

        private void AssignableDevice_Click(object sender, RoutedEventArgs e)
        {

        }

        private void RemAllDevice_Click(object sender, RoutedEventArgs e)
        {

        }

        private void HyperVRefresh_Click(object sender, RoutedEventArgs e)
        {
            StatusBarChangeBehaviour(true);

            RefreshVMs();

            StatusBarChangeBehaviour(false);
        }

        private void AboutBox_Click(object sender, RoutedEventArgs e)
        {
            About aboutbox = new About();
            aboutbox.ShowDialog();
        }

        private void ChangeCacheTypes_Click(object sender, RoutedEventArgs e)
        {

        }

        /*
         * Non-button methods
         */

        private void RefreshVMs()
        {
            try
            {
                var devices = this.wmi.getManagementObjectCollection("Msvm_ComputerSystem", "root\\virtualization\\v2", "Caption, ElementName, Name, EnabledState");
                foreach (var vm in devices)
                {
                    if (vm["Caption"].Equals("Virtual Machine"))
                    {
                        string vmName = vm["ElementName"]?.ToString() ?? vm["Name"].ToString();
                        string vmStatus = vmStatusMap[int.Parse(vm["EnabledState"]?.ToString() ?? "0")];
                        VMList.Items.Add(new
                        {
                            VMName = vmName,
                            VMStatus = vmStatus
                        });
                    }
                }
            }
            catch (ManagementException e)
            {
                MessageBox.Show(e.ToString(), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
        }

        private void StatusBarChangeBehaviour(bool isRefresh)
        {
            BottomProgressBarStatus.IsIndeterminate = isRefresh;
            if (isRefresh)
            {
                BottomLabelStatus.Text = "Refreshing...";
            }
            else
            {
                BottomLabelStatus.Text = "Done";
            }
        }
    }
}
