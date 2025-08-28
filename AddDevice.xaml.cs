using DDAGUI.WMIProperties;
using System;
using System.Management;
using System.Runtime.InteropServices;
using System.Windows;

namespace DDAGUI
{
    public partial class AddDevice : Window
    {
        protected string deviceId;
        protected MachineMethods machine;

        public AddDevice(MachineMethods machine)
        {
            this.machine = machine;
            machine.Connect("root\\cimv2");
            InitializeComponent();
        }

        /*
         * Button and UI methods behaviour....
         */

        private void AddDeviceButton_Click(object sender, RoutedEventArgs e)
        {
            if (DeviceList.SelectedItem != null)
            {
                deviceId = DeviceList.SelectedItem.GetType().GetProperty("DeviceId").GetValue(DeviceList.SelectedItem, null).ToString();
                DialogResult = true;
            }
            else
            {
                MessageBox.Show("Please select a device to add.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void AddDeviceCloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /*
         * Non-button methods
         */

        public string GetDeviceId()
        {
            UpdateDevices();
            ShowDialog();

            if (!(deviceId == null || deviceId == string.Empty))
            {
                deviceId = deviceId.Trim();
            }
            else
            {
                throw new NullReferenceException("[DDA Error] The device either has no valid ID or has disconnected.");
            }

            return deviceId;
        }

        private void UpdateDevices()
        {
            DeviceList.Items.Clear();

            try
            {
                foreach (var device in machine.GetObjects("Win32_PnPEntity", "Status, PNPClass, Name, DeviceID"))
                {
                    string deviceStatus = device["Status"]?.ToString() ?? "Unknown";
                    string deviceType = device["PNPClass"]?.ToString() ?? "Unknown";
                    string deviceName = device["Name"]?.ToString() ?? "Unknown";
                    string deviceId = device["DeviceID"]?.ToString() ?? "Unknown";

                    if (deviceStatus != "OK" || !deviceId.StartsWith("PCI"))
                    {
                        continue;
                    }

                    DeviceList.Items.Add(new
                    {
                        DeviceStatus = deviceStatus,
                        DeviceType = deviceType,
                        DeviceName = deviceName,
                        DeviceId = deviceId
                    });
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
    }
}
