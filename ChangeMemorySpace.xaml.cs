using System.Windows;

/*
 * Primary namespace for HyperV-DPD application
 *  It contains the main window and all related methods for the core application
 */
namespace TheFlightSims.HyperVDPD
{
    /*
     * Change Memory Space Window
     */
    public partial class ChangeMemorySpace : Window
    {
        ////////////////////////////////////////////////////////////////
        /// Global Properties and Constructors Region
        ///     This region contains global properties and constructors 
        ///     for the MachineMethods class.
        ////////////////////////////////////////////////////////////////
        
        /*
         * Global properties
         *  memRange as tuple of low and high memory
         */
        protected (ulong lowMem, ulong highMem) memRange;

        // Constructor of the Change Memory space window
        public ChangeMemorySpace(string vmName)
        {
            Title += $" for {vmName}";
            InitializeComponent();
        }

        ////////////////////////////////////////////////////////////////
        /// User Action Methods Region
        ///     This region contains methods that handle user actions.
        ///     For example, button clicks, changes in order.
        ////////////////////////////////////////////////////////////////

        /*
         * Actions for button "Confirm"
         */
        private void Confirm_Button(object sender, RoutedEventArgs e)
        {
            // Try to parse number in the box
            if (ulong.TryParse(LowMemory_TextBox.Text, out ulong lowMemCompare) && ulong.TryParse(HighMemory_TextBox.Text, out ulong highMemCompare))
            {
                // Check for low MMIO and high MMIO. If valid, return the value
                if ((lowMemCompare >= 128 && lowMemCompare <= 3584) && (highMemCompare >= 4096 && highMemCompare <= (ulong.MaxValue - 2)))
                {
                    memRange.lowMem = lowMemCompare;
                    memRange.highMem = highMemCompare;
                    DialogResult = true;
                }
                else
                {
                    // Notify user for invalid valid input
                    _ = MessageBox.Show(
                        "Make sure the Low MMIO Gap is in range (128, 3584) and High MMIO is larger than 4096",
                        "Warning",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning
                    );
                }
            }
            else
            {
                _ = MessageBox.Show(
                    "Please enter a positive integer in both box.",
                    "Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
            }
        }

        /*
         * Action for button "Cancel"
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
         * Return Value
         */
        public (ulong, ulong) ReturnValue()
        {
            _ = ShowDialog();

            return memRange;
        }
    }
}
