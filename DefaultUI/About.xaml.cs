using System.Windows;

/* The Default UI includes
 * 1. ConnectForm.xaml - A form to connect to a local or remote computer
 * 2. ExceptionView.xaml - A window to handle exception messages
 * 3. About.xaml - A window to display information about the application
 */
namespace TheFlightSims.HyperVDPD.DefaultUI
{
    /*
     * About Form class
     * It is used to display the help info & Legal notice
     */
    public partial class About : Window
    {
        // Constructor of the menu
        public About()
        {
            InitializeComponent();
        }
    }
}
