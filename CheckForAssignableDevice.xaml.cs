using DDAGUI.WMIProperties;
using System;
using System.Collections.Generic;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;

namespace DDAGUI
{
    public partial class CheckForAssignableDevice : Window
    {
        protected MachineMethods machine;

        public CheckForAssignableDevice(MachineMethods machine)
        {
            this.machine = machine;
            
            InitializeComponent();
        }

        /*
         * Button and UI methods behaviour....
         */
        private async void Check_Button(object sender, RoutedEventArgs e)
        {
            await CheckDevices();
        }

        private void Cancel_Button(object sender, RoutedEventArgs e) 
        {
            Close();
        }

        /*
         * Non-button methods
         */
        private async Task CheckDevices()
        {

            AssignableDevice_ListView.Items.Clear();

            try
            {
                await Task.Run(() => {
                    machine.Connect("root\\cimv2");

                    foreach (var device in machine.GetObjects("Win32_PnPEntity", "DeviceID, Caption, Status"))
                    {
                        if (device["DeviceID"].ToString().StartsWith("PCI\\"))
                        {
                            bool isAssignable = true;
                            string deviceId = device["DeviceID"].ToString();
                            string deviceName = device["Caption"].ToString();
                            HashSet<string> startingAddresses = new HashSet<string>();
                            string deviceNote = string.Empty;
                            double memoryGap = 0;

                            foreach (var actualPnPDevice in machine.GetObjects("Win32_PnPDevice", "SameElement, SystemElement"))
                            {
                                if (actualPnPDevice["SystemElement"].ToString().Contains(deviceId.Replace("\\", "\\\\")))
                                {
                                    isAssignable = true;
                                    break;
                                }
                                else
                                {
                                    isAssignable = false;
                                }
                            }

                            if (!isAssignable)
                            {
                                deviceNote += ((deviceNote.Length == 0) ? "" : "\n") + "The PCI device is not support either Express Endpoint, Embedded Endpoint, or Legacy Express Endpoint.";
                            }

                            foreach (var deviceResource in machine.GetObjects("Win32_PNPAllocatedResource", "Antecedent, Dependent"))
                            {
                                if (deviceResource["Dependent"].ToString().Contains(deviceId.Replace("\\", "\\\\")))
                                {
                                    string addr = ((string[])deviceResource["Antecedent"].ToString().Split('='))[1].Replace("\"", "");
                                    startingAddresses.Add(addr);
                                }
                            }

                            if (!device["Status"].ToString().ToLower().Equals("ok"))
                            {
                                deviceNote += ((deviceNote.Length == 0) ? "" : "\n") + "Device is disabled. Please re-enable the device to let the program check the memory gap.";
                                isAssignable = false;
                            }

                            if (isAssignable)
                            {
                                foreach (var deviceMem in machine.GetObjects("Win32_DeviceMemoryAddress", "StartingAddress, EndingAddress"))
                                {
                                    if(startingAddresses.Contains(deviceMem["StartingAddress"].ToString()))
                                    {
                                        double startingRange = Double.Parse(deviceMem["StartingAddress"].ToString());
                                        double endingRange = Double.Parse(deviceMem["EndingAddress"].ToString());
                                        memoryGap += Math.Ceiling((endingRange - startingRange) / 1048576.0);
                                    }
                                }
                            }

                            Dispatcher.Invoke(() =>
                            {
                                AssignableDevice_ListView.Items.Add(new
                                {
                                    DeviceId = deviceId,
                                    DeviceName = deviceName,
                                    DeviceGap = (!isAssignable) ? WMIDefaultValues.notAvailable : $"{memoryGap} MB",
                                    DeviceNote = deviceNote
                                });
                            });
                        }
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
    }
}
