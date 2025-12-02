using System;
using System.Linq;
using System.Management;
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
     * Add Device Window class
     * It is used to add device into the selected VM
     */
    public partial class AddDevice : Window
    {
        ////////////////////////////////////////////////////////////////
        /// Global Properties and Constructors Region
        ///     This region contains global properties and constructors 
        ///     for the MachineMethods class.
        ////////////////////////////////////////////////////////////////

        /*
         *  Global properties
         *      deviceId: Device ID of the selected device
         *      machine: WMI namespace of the machine
         */

        protected string deviceId;
        protected MachineMethods machine;

        /*
         * Constructor of the MainWindow class
         * Initializes the components and sets up event handlers
         */
        public AddDevice(MachineMethods machine)
        {
            this.machine = machine;
            machine.Connect("root\\cimv2");
            InitializeComponent();
        }


        ////////////////////////////////////////////////////////////////
        /// User Action Methods Region
        ///     This region contains methods that handle user actions.
        ///     For example, button clicks, changes in order.
        ////////////////////////////////////////////////////////////////

        /*
         * Actions for button "Add Device"
         */
        private void AddDeviceButton_Click(object sender, RoutedEventArgs e)
        {
            // If the selected item is not null (nothing is selected), proceed
            if (DeviceList.SelectedItem != null)
            {
                deviceId = DeviceList.SelectedItem.GetType().GetProperty("DeviceId").GetValue(DeviceList.SelectedItem, null).ToString();
                DialogResult = true;
            }
            else
            {
                // Notify user if nothing is selected
                _ = MessageBox.Show(
                    "Please select a device to add.",
                    "Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
            }
        }

        /*
         * Actions for button "Close"
         */
        private void AddDeviceCloseButton_Click(object sender, RoutedEventArgs e)
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
         * Start getting the device ID
         */
        public string GetDeviceId()
        {

            UpdateDevices();
            _ = ShowDialog();

            // Return the device ID back to the main windows
            return deviceId;
        }

        /*
         * Update device list
         */
        private async void UpdateDevices()
        {
            // Clear the device list
            DeviceList.Items.Clear();

            try
            {
                // Loop for devices in Win32_PnPEntity
                foreach (ManagementObject device in machine.GetObjects("Win32_PnPEntity", "Status, PNPClass, Name, DeviceID").Cast<ManagementObject>())
                {
                    string deviceStatus = device["Status"]?.ToString() ?? "Unknown";
                    string deviceType = device["PNPClass"]?.ToString() ?? "Unknown";
                    string deviceName = device["Name"]?.ToString() ?? "Unknown";
                    string deviceId = device["DeviceID"]?.ToString() ?? "Unknown";

                    // If the device is not PCI, skip to the next device
                    if (!deviceId.StartsWith("PCI"))
                    {
                        device.Dispose();
                        continue;
                    }

                    _ = DeviceList.Items.Add(new
                    {
                        DeviceStatus = deviceStatus,
                        DeviceType = deviceType,
                        DeviceName = deviceName,
                        DeviceId = deviceId
                    });

                    device.Dispose();
                }
            }
            catch (Exception ex)
            {
                (new ExceptionView()).HandleException(ex);
            }
        }
    }
}
