using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Windows;

/* The Default UI includes
 * 1. ConnectForm.xaml - A form to connect to a local or remote computer
 * 2. ExceptionView.xaml - A window to handle exception messages
 */
namespace TheFlightSims.HyperVDPD.DefaultUI
{
    /*
     * Exception View class
     * It is used to handle exception messages
     * 
     * Excepted behaviour:
     *  1. An exception is passed to the HandleException method
     *  2. If the exception is of an allowed type, show the exception message 
     *      and give the user the option to ignore it; else, force close the application
     */
    public partial class ExceptionView : Window
    {
        ////////////////////////////////////////////////////////////////
        /// Global Properties and Constructors Region
        ///     This region contains global properties and constructors 
        ///     for the MachineMethods class.
        ////////////////////////////////////////////////////////////////

        /*
         * Global properties
         *  allowedExceptionTypes: A list of exception types that are allowed to be ignored
         */
        private readonly HashSet<Type> allowedExceptionTypes = new HashSet<Type>
        {
            typeof(UnauthorizedAccessException),
            typeof(COMException),
            typeof(ManagementException),
            typeof(NullReferenceException)
        };

        // Constructor of the Exception View class
        public ExceptionView()
        {
            InitializeComponent();
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
         * Click Ignore Exception Button - The "Ignore" button
         * Let user ignore the exception and continue using the application
         */
        private void Click_IgnoreException(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /*
         * Click Close Application Button - The "Close Application" button
         * Close the application in case of a critical or unhandled exception
         */
        private void Click_CloseApp(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        ////////////////////////////////////////////////////////////////
        /// User Action Methods Region
        ///     This region contains methods that handle user actions.
        ///     For example, button clicks, changes in order.
        ////////////////////////////////////////////////////////////////

        /*
         * Handle Exception Method
         * Get the exception info, and decide whether to allow the user to ignore it
         */
        public void HandleException(Exception e)
        {
            if (e == null)
            {
                return;
            }

            // Check if the exception type is in the allowed list
            bool isAllowed = this.allowedExceptionTypes.Any(t => t.IsInstanceOfType(e));

            if (!isAllowed)
            {
                ExceptionMessage_TextBlock.Text = "An unexpected error has occurred.\n" +
                    "You should close the application and check for the system";
                Button_Ignore.Visibility = Visibility.Collapsed;
            }

#if !DEBUG
            DetailedExceptionDetail_TextBox.Text = e.Message;
#else
            DetailedExceptionDetail_TextBox.Text = e.ToString();
#endif

            ShowDialog();
        }
    }
}
