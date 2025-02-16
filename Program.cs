using System;
using System.Windows.Forms;

namespace DDA_GUI
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new DDA());
        }
    }
}
