using System.Windows;

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
        }

        private void WMIConnect_Click(object sender, RoutedEventArgs e)
        {
            ConnectForm connectForm = new ConnectForm();
            connectForm.ShowsNavigationUI = true;
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

        private void HyperVRefresh_Click(object sender, RoutedEventArgs e)
        {

        }

        private void AboutBox_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ChangeCacheTypes_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}