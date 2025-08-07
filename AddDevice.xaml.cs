using System;
using System.Windows;

using DDAGUI.WMIProperties;


namespace DDAGUI
{
    public partial class AddDevice : Window
    {
        protected string deviceId;
        protected WMIWrapper wmi;

        public AddDevice(WMIWrapper wmi)
        {
            this.wmi = wmi;
            InitializeComponent();
        }

        /*
         * Button and UI methods behaviour....
         */

        private void AddDeviceButton_Click(object sender, RoutedEventArgs e)
        {
            if (DeviceList.SelectedItem != null)
            {
                this.deviceId = DeviceList.SelectedItem.GetType().GetProperty("DeviceId").GetValue(DeviceList.SelectedItem, null).ToString();
                this.DialogResult = true;
            }
            else
            {
                MessageBox.Show("Please select a device to add.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void AddDeviceCloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /*
         * Non-button methods
         */

        public string GetDeviceId()
        {
            UpdateDevices(this.wmi);
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

        private void UpdateDevices(WMIWrapper wmi)
        {
            DeviceList.Items.Clear();
            foreach (var device in wmi.GetManagementObjectCollection("Win32_PnPEntity", "root\\cimv2", "Status, PNPClass, Name, DeviceID"))
            {
                string deviceStatus = device["Status"]?.ToString() ?? "Unknown";
                string deviceType = device["PNPClass"]?.ToString() ?? "Unknown";
                string deviceName = device["Name"]?.ToString() ?? "Unknown";
                string deviceId = device["DeviceID"]?.ToString() ?? "Unknown";

                if (deviceStatus != "OK" || !deviceId.StartsWith("PCI")) {
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
    }
}
