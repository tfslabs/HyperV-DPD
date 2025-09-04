using TheFlightSims.HyperVDPD.WMIProperties;
using System;
using System.Linq;
using System.Management;
using System.Windows;

namespace TheFlightSims.HyperVDPD
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
                foreach (ManagementObject srv in machine.GetObjects("Win32_Service", "Name, Caption, State").Cast<ManagementObject>())
                {
                    if (srv["Name"] == null || !WMIDefaultValues.serviceNames.Contains(srv["Name"].ToString()))
                    {
                        srv.Dispose();
                        continue;
                    }

                    ListHVSrv.Items.Add(new
                    {
                        ServiceName = srv["Caption"]?.ToString() ?? "Unknown",
                        ServiceStatus = srv["State"]?.ToString() ?? "Unknown"
                    });

                    srv.Dispose();
                }
            }
            catch (Exception ex)
            {
                WMIDefaultValues.HandleException(ex, machine.GetComputerName());
            }
        }
    }
}
