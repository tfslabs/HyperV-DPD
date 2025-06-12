using System;
using System.Windows;
using System.Management;
using Microsoft.Win32;

namespace DDAGUI
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            /*
             * TO DO
             * 
             *
             * 
             * Registry to get the nformation about the Windows version
             * - Get the Windows version and build number
             * - Compare if the build number is greater than or equal to 19041 (Windows 10 20H2)
             * - If the build number is less than 19041, show a message box that Hyper-V with DDA 
             *   is not supported and exit the application on Release mode, or continue in Debug mode
             * - Also in Debug mode, show the Windows version and build number in the title bar
             * - Also handle for if in any case the registry key is not found (including running on a non-Windows OS)
             * 
             * Get if Hyper-V is enabled and running
             * - If Hyper-V is not enabled, show a message box that Hyper-V is not enabled and exit the application 
             * on Release mode, or continue in Debug mode
             * - Also handle for if in any case the registry key is not found (including running on a non-Windows OS)
             * 
             */

#if DEBUG
            this.Title += " (Debugging is enabled)";
#endif

            //////////////////////////////////////////////////////////
            //// Check Windows Version and Build Number
#if !DEBUG
            try
            {
#endif
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
                int buildNumber = Convert.ToInt32(key.GetValue("CurrentBuild"));

                if (buildNumber >= 19041) {

                    this.Title += " - Windows Build " + buildNumber.ToString();
                } 
                else
                {
                    if (MessageBox.Show("This application requires Windows 10 version 19041 or later\nDo you wish to continue", 
                        "Unsupported Windows Version", 
                        MessageBoxButton.YesNo, 
                        MessageBoxImage.Warning) 
                        == MessageBoxResult.No)
                    {
                        Application.Current.Shutdown();
                    }

                   Application.Current.Shutdown();
                }

#if !DEBUG
            }
            catch (Exception regError)
            {
                MessageBox.Show(regError.Message);
            }
#endif
            //// End of Windows Version and Build Number Check
            //////////////////////////////////////////////////////////

            //////////////////////////////////////////////////////////
            //// Check if Hyper-V is enabled
#if !DEBUG
            try
            {
#endif
                bool isHyperVEnabled = false;
                using (var searcher = new ManagementObjectSearcher("root\\virtualization", "SELECT * FROM Msvm_ComputerSystem"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        isHyperVEnabled = true;
                        break;
                    }
                }
                
                if (!isHyperVEnabled)
                {
                    MessageBox.Show("Hyper-V is not enabled.");
                    Application.Current.Shutdown();
                }
#if !DEBUG
            }
            catch (Exception ex) {
                MessageBox.Show("An error occurred while checking Hyper-V status: " + ex.HResult);
            }
#endif
            //// End of Hyper-V Enabled Check
            //////////////////////////////////////////////////////////
        }

        private void WMIConnect_Click(object sender, RoutedEventArgs e)
        {

        }

        private void QuitMainWindow_Click(object sender, RoutedEventArgs e)
        {

        }

        private void AddDevice_Click(object sender, RoutedEventArgs e)
        {

        }

        private void RemDevice_Click(object sender, RoutedEventArgs e)
        {

        }

        private void CopyDevAddress_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ChangeMemLocation_Click(object sender, RoutedEventArgs e)
        {

        }

        private void IsGuestControlledCacheTypes_Click(object sender, RoutedEventArgs e)
        {

        }

        private void HyperVServiceStatus_Click(object sender, RoutedEventArgs e)
        {

        }

        private void AssignableDevice_Click(object sender, RoutedEventArgs e)
        {

        }

        private void RemAllDevice_Click(object sender, RoutedEventArgs e)
        {

        }

        private void HyperVRefresh_Click (object sender, RoutedEventArgs e)
        {

        }

        private void AboutBox_Click (object sender, RoutedEventArgs e)
        {

        }

        private void ChangeCacheTypes_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
