using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Threading.Tasks;
using System.Windows;
using TheFlightSims.HyperVDPD.WMIProperties;

namespace TheFlightSims.HyperVDPD
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
                await Task.Run(() =>
                {
                    machine.Connect("root\\cimv2");

                    foreach (ManagementObject device in machine.GetObjects("Win32_PnPEntity", "DeviceID, Caption, Status").Cast<ManagementObject>())
                    {
                        string deviceId = device["DeviceID"]?.ToString();
                        string devCaption = device["Caption"]?.ToString();
                        string devStatus = device["Status"]?.ToString();

                        if (deviceId == null || devCaption == null || devStatus == null)
                        {
                            device.Dispose();
                            continue;
                        }

                        if (deviceId.StartsWith("PCI\\"))
                        {
                            bool isAssignable = true;
                            HashSet<string> startingAddresses = new HashSet<string>();
                            string deviceNote = string.Empty;
                            double memoryGap = 0;

                            foreach (ManagementObject actualPnPDevice in machine.GetObjects("Win32_PnPDevice", "SameElement, SystemElement").Cast<ManagementObject>())
                            {
                                string actualPnpDevString = actualPnPDevice["SystemElement"]?.ToString();

                                if (actualPnpDevString == null)
                                {
                                    isAssignable = false;
                                    actualPnPDevice.Dispose();
                                    continue;
                                }

                                if (actualPnpDevString.Contains(deviceId.Replace("\\", "\\\\")))
                                {
                                    isAssignable = true;
                                    break;
                                }
                                else
                                {
                                    isAssignable = false;
                                }

                                actualPnPDevice.Dispose();
                            }

                            if (!isAssignable)
                            {
                                deviceNote += ((deviceNote.Length == 0) ? "" : "\n") + "The PCI device is not support either Express Endpoint, Embedded Endpoint, or Legacy Express Endpoint.";
                            }

                            foreach (ManagementObject deviceResource in machine.GetObjects("Win32_PNPAllocatedResource", "Antecedent, Dependent").Cast<ManagementObject>())
                            {
                                string devResDependent = deviceResource["Dependent"]?.ToString();

                                if (devResDependent == null)
                                {
                                    deviceResource.Dispose();
                                    continue;
                                }

                                if (devResDependent.Contains(deviceId.Replace("\\", "\\\\")))
                                {
                                    startingAddresses.Add((deviceResource["Antecedent"].ToString().Split('=')[1]).Replace("\"", ""));
                                }

                                deviceResource.Dispose();
                            }

                            if (!devStatus.ToLower().Equals("ok"))
                            {
                                deviceNote += ((deviceNote.Length == 0) ? "" : "\n") + "Device is disabled. Please re-enable the device to let the program check the memory gap.";
                                isAssignable = false;
                            }

                            if (isAssignable)
                            {
                                foreach (ManagementObject deviceMem in machine.GetObjects("Win32_DeviceMemoryAddress", "StartingAddress, EndingAddress").Cast<ManagementObject>())
                                {
                                    string startAddrString = deviceMem["StartingAddress"]?.ToString();
                                    string endAddrString = deviceMem["EndingAddress"]?.ToString();

                                    if (startAddrString == null || endAddrString == null)
                                    {
                                        deviceMem.Dispose();
                                        continue;
                                    }

                                    if (startingAddresses.Contains(startAddrString))
                                    {
                                        memoryGap += Math.Ceiling((Double.Parse(endAddrString) - Double.Parse(startAddrString)) / 1048576.0);
                                    }

                                    deviceMem.Dispose();
                                }
                            }

                            Dispatcher.Invoke(() =>
                            {
                                AssignableDevice_ListView.Items.Add(new
                                {
                                    DeviceId = deviceId,
                                    DeviceName = devCaption,
                                    DeviceGap = (!isAssignable) ? WMIDefaultValues.notAvailable : $"{memoryGap} MB",
                                    DeviceNote = deviceNote
                                });
                            });
                        }

                        device.Dispose();
                    }
                });
            }
            catch (Exception ex)
            {
                WMIDefaultValues.HandleException(ex, machine.GetComputerName());
            }
        }
    }
}
