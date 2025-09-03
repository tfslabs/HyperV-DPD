using System;
using System.Collections.Generic;
using System.Windows;

#if !DEBUG
using System.Runtime.InteropServices;
using System.Management;
#endif

namespace DDAGUI.WMIProperties
{
    public static class WMIDefaultValues
    {
        public static string notAvailable = "N/A";

        public static readonly Dictionary<UInt16, string> vmStatusMap = new Dictionary<UInt16, string>
        {
            {0,  "Unknown" },
            {1,  "Other" },
            {2,  "Running" },
            {3,  "Stopped" },
            {4,  "Shutting down" },
            {5,  "Not applicable" },
            {6,  "Enabled but Offline" },
            {7,  "In Test" },
            {8,  "Degraded" },
            {9,  "Quiesce" },
            {10, "Starting" }
        };

        public static readonly HashSet<string> serviceNames = new HashSet<string>()
        {
            "HvHost",
            "vmickvpexchange",
            "gcs",
            "vmicguestinterface",
            "vmicshutdown",
            "vmicheartbeat",
            "vmcompute",
            "vmicvmsession",
            "vmicrdv",
            "vmictimesync",
            "vmms",
            "vmicvss"
        };

        public static void HandleException(Exception ex, string machineName)
        {
#if DEBUG
            MessageBox.Show(
                ex.ToString(),
                $"Error on {machineName}",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            );
#else
            string message = String.Empty;
            if (ex is UnauthorizedAccessException)
            {
                message = $"Failed to catch the Authenticate with {machineName}: {ex.Message}";
            }
            else if (ex is COMException)
            {
                message = $"Failed to reach {machineName}: {ex.Message}";
            }
            else if (ex is ManagementException)
            {
                message = $"Failed to catch the Management Method: {ex.Message}";
            }
            else if (ex is NullReferenceException)
            {
                message = "No input was made";
            }
            else
            {
                message = ex.Message;
            }

            MessageBox.Show(message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
        }
    }
}
