using System.Security;
using System.Windows;

namespace TheFlightSims.HyperVDPD
{
    public partial class ConnectForm : Window
    {
        /*
         *  Global properties
         */
        private bool isConnectLocally;

        public ConnectForm()
        {
            isConnectLocally = false;
            InitializeComponent();
        }

        /* 
         * Button-based methods
         */

        private void ConnetToExternal_Button(object sender, RoutedEventArgs e)
        {
            if (ComputerName_TextBox.Text.Length == 0 || UserName_TextBox.Text.Length == 0)
            {
                MessageBox.Show(
                    "Please don't leave computer name and user name empty",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
            }
            else
            {
                DialogResult = true;
            }
        }

        private void ConnectToLocal_Button(object sender, RoutedEventArgs e)
        {
            isConnectLocally = true;
            DialogResult = true;
        }

        private void Cancel_Button(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /*
         * Non-button methods
         */
        public (string, string, SecureString) ReturnValue()
        {
            ShowDialog();

            string computerName = (isConnectLocally) ? "localhost" : ComputerName_TextBox.Text;
            string userName = (isConnectLocally) ? "" : UserName_TextBox.Text;
            SecureString password = new SecureString();

            if (!isConnectLocally)
            {
                foreach (char c in Password_TextBox.Password)
                {
                    password.AppendChar(c);
                }
            }

            return (computerName, userName, password);
        }
    }
}
