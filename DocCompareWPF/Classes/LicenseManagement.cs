using System;
using System.Collections.Generic;
using System.Management;
using System.Text;

namespace DocCompareWPF.Classes
{
    class LicenseManagement
    {
        public string UUID;

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
            }
            catch (Exception ex)
            {
                ErrorHandling.ReportException(ex);
            }
        }
    }
}
