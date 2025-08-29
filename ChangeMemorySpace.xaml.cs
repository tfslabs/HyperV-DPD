using System.Windows;

namespace DDAGUI
{
    public partial class ChangeMemorySpace : Window
    {
        /*
         * Global properties
         */
        protected (int lowMem, int highMem) memRange;

        public ChangeMemorySpace(string vmName)
        {
            Title += $" for {vmName}";
            InitializeComponent();
        }

        /*
         * Button-based methods
         */
        private void Confirm_Button(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(LowMemory_TextBox.Text, out int lowMemCompare) && int.TryParse(HighMemory_TextBox.Text, out int highMemCompare))
            {
                if ((lowMemCompare < highMemCompare) && (lowMemCompare > 0 && highMemCompare > 0))
                {
                    memRange.lowMem = lowMemCompare;
                    memRange.highMem = highMemCompare;
                    DialogResult = true;
                }
                else
                {
                    MessageBox.Show("Make sure the Low Memory is lower than High Memory and both are larger than 0.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                MessageBox.Show("Please enter a positive integer in both box.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Cancel_Button(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /*
         * Non-button methods
         */
        public (int, int) ReturnValue()
        {
            ShowDialog();
            
            return memRange;
        }
    }
}
