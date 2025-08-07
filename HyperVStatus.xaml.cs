using System.Collections.Generic;
using System.Windows;

using DDAGUI.WMIProperties;

namespace DDAGUI
{
    public partial class HyperVStatus : Window
    {
        protected WMIWrapper wmi;
        protected HashSet<string> serviceNames = new HashSet<string>()
        {
            "HvHost",
            "vmickvpexchange",
            "gcs",
            "vmicguestinterface",
            "vmicshutdown",
            "vmicheartbeat",
            "vmcompute",
            "vmicvmsession",
            "vmicrdv",
            "vmictimesync",
            "vmms",
            "vmicvss"
        };

        public HyperVStatus(WMIWrapper wmi)
        {
            this.wmi = wmi;
            InitializeComponent();
            this.RefreshServices();
        }

        /*
         * Button and UI methods behaviour....
         */

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            ListHVSrv.Items.Clear();
            this.RefreshServices();
        }

        /*
         * Non-button methods
         */

        private void RefreshServices()
        {
            foreach (var srv in this.wmi.GetManagementObjectCollection("Win32_Service", "root\\cimv2"))
            {
                if (srv["Name"] == null || !serviceNames.Contains(srv["Name"].ToString()))
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
    }
}
