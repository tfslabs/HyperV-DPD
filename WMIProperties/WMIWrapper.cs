using System.Management;

namespace DDAGUI.WMIProperties
{
    public class WMIWrapper
    {
        protected string computerName;

        public WMIWrapper(string computerName)
        {
            this.computerName = computerName;
        }

        public void SetComputerName(string computerName)
        {
            this.computerName = computerName;
        }

        public ManagementObjectCollection GetManagementObjectCollection(string className, string nameSpace, string fields = "*")
        {
            ManagementObjectSearcher searcher;

            try
            {
                ManagementScope scope = new ManagementScope($"\\\\{this.computerName}\\{nameSpace}");
                ObjectQuery query = new ObjectQuery($"SELECT {fields} FROM {className}");

                scope.Connect();
                searcher = new ManagementObjectSearcher(scope, query);
            }
            catch (ManagementException e)
            {
#if DEBUG
                throw new ManagementException("Error of {GetManagementObjectCollection}:\n" + e.ToString());
#else
                throw new ManagementException($"Failed to connect to WMI namespace {nameSpace} on {ComputerName}: {e.Message}");
#endif
            }
            finally
            {
                //
            }

            return searcher?.Get();
        }
    }
}
