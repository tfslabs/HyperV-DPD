using System;
using System.Collections.Generic;
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
     * Hyper-V Services Status Window
     */
    public partial class HyperVStatus : Window
    {
        ////////////////////////////////////////////////////////////////
        /// Global Properties and Constructors Region
        ///     This region contains global properties and constructors 
        ///     for the MachineMethods class.
        ////////////////////////////////////////////////////////////////
        
        
        protected MachineMethods machine;
        public static readonly HashSet<string> serviceNames = new HashSet<string>()
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
                    if (srv["Name"] == null || !serviceNames.Contains(srv["Name"].ToString()))
                    {
                        srv.Dispose();
                        continue;
                    }

                    _ = ListHVSrv.Items.Add(new
                    {
                        ServiceName = srv["Caption"]?.ToString() ?? "Unknown",
                        ServiceStatus = srv["State"]?.ToString() ?? "Unknown"
                    });

                    srv.Dispose();
                }
            }
            catch (Exception ex)
            {
                (new ExceptionView()).HandleException(ex);
            }
        }
    }
}
