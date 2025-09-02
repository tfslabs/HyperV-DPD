using System;
using System.Windows;

namespace DDAGUI
{
    public partial class ChangeMemorySpace : Window
    {
        /*
         * Global properties
         */
        protected (UInt64 lowMem, UInt64 highMem) memRange;

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
            if (UInt64.TryParse(LowMemory_TextBox.Text, out UInt64 lowMemCompare) && UInt64.TryParse(HighMemory_TextBox.Text, out UInt64 highMemCompare))
            {
                if ((lowMemCompare >= 128 && lowMemCompare <= 3584) && (highMemCompare >= 4096 && highMemCompare <= (UInt64.MaxValue - 2)))
                {
                    memRange.lowMem = (UInt64)(lowMemCompare);
                    memRange.highMem = (UInt64)(highMemCompare);
                    DialogResult = true;
                }
                else
                {
                    MessageBox.Show("Make sure the Low MMIO Gap is in range (128, 3584) and High MMIO is larger than 4096", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
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
        public (UInt64, UInt64) ReturnValue()
        {
            ShowDialog();

            return memRange;
        }
    }
}
