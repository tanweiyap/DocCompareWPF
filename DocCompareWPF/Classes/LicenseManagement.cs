using System;
using System.Management;

namespace DocCompareWPF.Classes
{
    internal class LicenseManagement
    {
        public enum LicenseTypes
        {
            TRIAL,
            ANNUAL_SUBSCRIPTION,
            DEVELOPMENT,
        };

        private string UUID;
        private LicenseTypes licenseType;        

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

        public string GetUUID()
        {
            return UUID;
        }

        public LicenseTypes GetLicenseTypes()
        {
            return licenseType;
        }

        public int ActivateLincense(string userEmail, string licKey)
        {
            if (Helper.IsValidEmail(userEmail))
            {
            }
            else
            {
                return -1;
            }

            return 0;
        }
    }
}