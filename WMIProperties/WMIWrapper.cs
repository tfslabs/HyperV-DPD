using System.Management;

namespace DDAGUI
{
    public class WMIWrapper
    {
        protected string ComputerName;

        public WMIWrapper(string computerName)
        {
            this.ComputerName = computerName;
        }

        public ManagementObjectCollection getManagementObjectCollection(string className, string nameSpace, string fields = "*")
        {
            ManagementObjectSearcher searcher;

            try
            {
                ManagementScope scope = new ManagementScope($"\\\\{this.ComputerName}\\{nameSpace}");
                ObjectQuery query = new ObjectQuery($"SELECT {fields} FROM {className}");

                scope.Connect();
                searcher = new ManagementObjectSearcher(scope, query);
            }
            catch (ManagementException e)
            {
#if DEBUG
                throw new ManagementException("Error of {getManagementObjectCollection}:\n" + e.ToString());
#else
                throw new ManagementException($"Failed to connect to WMI namespace {nameSpace} on {ComputerName}: {e.Message}");
#endif
            }
            finally
            {
                //
            }

            return (searcher != null) ? searcher.Get() : null;
        }
    }
}
