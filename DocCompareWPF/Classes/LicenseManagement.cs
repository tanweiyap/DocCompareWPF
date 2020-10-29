using System;
using System.Drawing.Text;
using System.Management;
using System.Net.Http
    ;
using System.Threading.Tasks;

namespace DocCompareWPF.Classes
{
    internal class LicenseManagement
    {
        private static readonly HttpClient client = new HttpClient();
        //private static readonly string LocalDirectory = Directory.GetCurrentDirectory();
        private static readonly string serverAddress = "http://licserver.portmap.host:46928/";

        private LicenseTypes licenseType;

        private string UUID;

        public LicenseManagement()
        {
            try
            {
                string ComputerName = "localhost";
                ManagementScope Scope;
                Scope = new ManagementScope(String.Format("\\\\{0}\\root\\CIMV2", ComputerName), null);
                Scope.Connect();
                ObjectQuery Query = new ObjectQuery("SELECT UUID FROM Win32_ComputerSystemProduct");
                ManagementObjectSearcher Searcher = new ManagementObjectSearcher(Scope, Query);

                foreach (ManagementObject WmiObject in Searcher.Get())
                {
                    UUID = WmiObject["UUID"].ToString();// String
                }

                // development mode
                licenseType = LicenseTypes.DEVELOPMENT;
            }
            catch (Exception ex)
            {
                ErrorHandling.ReportException(ex);
            }
        }

        public enum LicenseTypes
        {
            TRIAL,
            ANNUAL_SUBSCRIPTION,
            DEVELOPMENT,
        };

        public int ActivateLincense(string userEmail, string licKey)
        {
            if (Helper.IsValidEmail(userEmail))
            {
                ContactServer(serverAddress + "status");
            }
            else
            {
                return -1;
            }

            return 0;
        }

        public async Task ContactServer(string url)
        {
            var res = await client.GetAsync(serverAddress + "status");
        }

        public LicenseTypes GetLicenseTypes()
        {
            return licenseType;
        }

        public string GetUUID()
        {
            return UUID;
        }
    }
}