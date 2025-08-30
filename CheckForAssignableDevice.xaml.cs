using DDAGUI.WMIProperties;
using System.Windows;

namespace DDAGUI
{
    public partial class CheckForAssignableDevice : Window
    {
        protected MachineMethods machine;

        public CheckForAssignableDevice(MachineMethods machine)
        {
            this.machine = machine;
            InitializeComponent();
        }
    }
}
