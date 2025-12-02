using System.Security;
using System.Windows;

/* The Default UI includes
 * 1. ConnectForm.xaml - A form to connect to a local or remote computer
 * 2. ExceptionView.xaml - A window to handle exception messages
 * 3. About.xaml - A window to display information about the application
 */
namespace TheFlightSims.HyperVDPD.DefaultUI
{
    /*
     * Connection Form class
     * It is used to get connection information from the user
     * Expected behaviour:
     *  1. User enter the computer name, username, and password to connect to a remote computer
     *  2. The application tries verify the connection information
     *  3. If what user enters is valid, the form closes and returns the information
     */
    public partial class ConnectForm : Window
    {

        ////////////////////////////////////////////////////////////////
        /// Global Properties and Constructors Region
        ///     This region contains global properties and constructors 
        ///     for the MachineMethods class.
        ////////////////////////////////////////////////////////////////

        /*
         *  Global properties
         *  isConnectLocally: A boolean to indicate whether the user wants to connect to the local computer
         */
        private bool isConnectLocally;

        // Constructor of the ConnectForm class
        public ConnectForm()
        {
            isConnectLocally = false;
            InitializeComponent();
        }

        ////////////////////////////////////////////////////////////////
        /// User Action Methods Region
        ///     This region contains methods that handle user actions.
        ///     For example, button clicks, changes in order.
        ////////////////////////////////////////////////////////////////

        /*
         * Connect to External Button Click Event - The "Connect" button
         * Validates the input fields and sets the DialogResult to true if valid
         */
        private void ConnetToExternal_Button(object sender, RoutedEventArgs e)
        {
            // Validate that none of the input fields are empty
            if (
                ComputerName_TextBox.Text.Length == 0 ||
                UserName_TextBox.Text.Length == 0 ||
                Password_TextBox.Password.Length == 0)
            {
                _ = MessageBox.Show("Please don't leave anything empty", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                // Close the form with a DialogResult of true
                DialogResult = true;
            }
        }

        /*
         * Connect to Local Button Click Event - The "Connect to Local" button
         * Sets the isConnectLocally flag to true and sets the DialogResult to true
         */
        private void ConnectToLocal_Button(object sender, RoutedEventArgs e)
        {
            isConnectLocally = true;
            DialogResult = true;
        }

        /*
         * Cancel Button Click Event - The "Cancel" button
         * Closes the form without setting a DialogResult
         */
        private void Cancel_Button(object sender, RoutedEventArgs e)
        {
            Close();
        }
        
        ////////////////////////////////////////////////////////////////
        /// Non-User Action Methods Region
        /// 
        /// This region contains methods that do not handle user actions.
        /// 
        /// Think about this is the back-end section.
        /// It should not be in a seperated class, because it directly interacts with the UI elements.
        ////////////////////////////////////////////////////////////////

        /*
         * Return Value Method
         * Returns a tuple containing the computer name, username, and password as a SecureString
         */
        public (string, string, SecureString) ReturnValue()
        {
            _ = ShowDialog();

            // Prepare the return values based on whether connecting locally or remotely
            string computerName = (isConnectLocally) ? "localhost" : ComputerName_TextBox.Text;
            string userName = (isConnectLocally) ? "" : UserName_TextBox.Text;
            SecureString password = new SecureString();

            // Only populate the password if connecting remotely
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
