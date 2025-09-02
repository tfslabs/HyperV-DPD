using DDAGUI.WMIProperties;

using System;
using System.Linq;
using System.Management;
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

            return deviceId;
        }

        private void UpdateDevices()
        {
            DeviceList.Items.Clear();

            try
            {
                foreach (ManagementObject device in machine.GetObjects("Win32_PnPEntity", "Status, PNPClass, Name, DeviceID").Cast<ManagementObject>())
                {
                    string deviceStatus = device["Status"]?.ToString() ?? "Unknown";
                    string deviceType = device["PNPClass"]?.ToString() ?? "Unknown";
                    string deviceName = device["Name"]?.ToString() ?? "Unknown";
                    string deviceId = device["DeviceID"]?.ToString() ?? "Unknown";

                    if (!deviceId.StartsWith("PCI"))
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
            catch (Exception ex)
            {
                WMIDefaultValues.HandleException(ex, machine.GetComputerName());
            }
        }
    }
}
