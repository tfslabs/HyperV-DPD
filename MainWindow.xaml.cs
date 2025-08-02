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

#if DEBUG
            this.Title += " (Debugging is enabled)";
#endif

            //////////////////////////////////////////////////////////
            //// Check Windows Version and Build Number
            try
            {
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

            }
            catch (Exception regError)
            {
#if DEBUG
                
#else
                MessageBox.Show(regError.Message);
#endif
            }
            //// End of Windows Version and Build Number Check
            //////////////////////////////////////////////////////////

            //////////////////////////////////////////////////////////
            //// Check if Hyper-V is enabled
            try
            {
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
            }
            catch (Exception ex) {
                MessageBox.Show("An error occurred while checking Hyper-V status: " + ex.HResult);
            }
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
