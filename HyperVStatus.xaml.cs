using DDAGUI.WMIProperties;
using System;
using System.Management;
using System.Runtime.InteropServices;
using System.Windows;

namespace DDAGUI
{
    public partial class HyperVStatus : Window
    {
        protected MachineMethods machine;


        public HyperVStatus(MachineMethods machine)
        {
            InitializeComponent();
            this.machine = machine;
            RefreshServices();
        }

        /*
         * Button and UI methods behaviour....
         */

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            ListHVSrv.Items.Clear();
            RefreshServices();
        }

        /*
         * Non-button methods
         */

        private void RefreshServices()
        {
            try
            {
                machine.Connect("root\\cimv2");
                foreach (var srv in machine.GetObjects("Win32_Service", "Name, Caption, State"))
                {
                    if (srv["Name"] == null || !WMIDefaultValues.serviceNames.Contains(srv["Name"].ToString()))
                    {
                        continue;
                    }

                    ListHVSrv.Items.Add(new
                    {
                        ServiceName = srv["Caption"]?.ToString() ?? "Unknown",
                        ServiceStatus = srv["State"]?.ToString() ?? "Unknown"
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
