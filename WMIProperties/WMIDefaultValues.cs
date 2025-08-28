using System.Collections.Generic;

namespace DDAGUI.WMIProperties
{
    public static class WMIDefaultValues
    {
        public static readonly Dictionary<int, string> vmStatusMap = new Dictionary<int, string>
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
    }
}
