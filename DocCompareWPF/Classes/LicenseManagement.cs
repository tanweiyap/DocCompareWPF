using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Management;
using System.Net.Http;
using System.Threading.Tasks;

namespace DocCompareWPF.Classes
{
    internal class LicenseManagement
    {
        private static readonly HttpClient client = new HttpClient() { Timeout = new TimeSpan(0,0,20)};

        //private static readonly string LocalDirectory = Directory.GetCurrentDirectory();
        private static readonly string serverAddress = "http://licserver.portmap.host:46928/";

        private LicenseTypes licenseType;

        private LicenseStatus licenseStatus;

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
                licenseStatus = LicenseStatus.ACTIVE;
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
            UNKNOWN,
        };

        public enum LicenseStatus
        {
            ACTIVE,
            INACTIVE,
        };

        public enum LicServerResponse
        {
            UNREACHABLE,
            KEY_MISMATCH,
            ACCOUNT_NOT_FOUND,
            OKAY,
        };

        public async Task<LicServerResponse> ActivateLincense(string userEmail, string licKey)
        {
            int status = await CheckServerStatus(serverAddress + "status");
            if (status == 0)
            {
                IDictionary<string, string> licDict = new Dictionary<string, string>
                    {
                        { "Email", userEmail},
                        { "LicKey", licKey},
                        { "UUID", UUID }
                    };

                var content = new FormUrlEncodedContent(licDict);

                HttpResponseMessage msg = await client.PostAsync(serverAddress + "activate", content);
                string readMsg = msg.Content.ReadAsStringAsync().Result;
                JObject resp = JsonConvert.DeserializeObject<JObject>(readMsg);

                if ((string)resp["CorrectKey"] == "true")
                {
                    licenseType = ParseLicTypes((string)resp["LicType"]);
                    licenseStatus = ParseLicStatus((string)resp["LicStatus"]);
                }
                else
                {
                    licenseType = LicenseTypes.UNKNOWN;
                    licenseStatus = LicenseStatus.INACTIVE;

                    if (ParseLicTypes((string)resp["LicType"]) == LicenseTypes.UNKNOWN)
                    {
                        return LicServerResponse.ACCOUNT_NOT_FOUND; //
                    }

                    return LicServerResponse.KEY_MISMATCH; // wrong key 
                }
            }
            else
            {
                return LicServerResponse.UNREACHABLE; // server offline
            }

            return LicServerResponse.OKAY;
        }

        public async Task<int> CheckServerStatus(string url)
        {
            try
            {
                string res = await client.GetStringAsync(url);

                if (res == "online")
                    return 0;
                else
                    return -1;
            }
            catch
            {
                ErrorHandling.ReportError("License server not reachabled", "-", "-");
                return -1;
            }
        }

        public LicenseTypes GetLicenseTypes()
        {
            return licenseType;
        }

        public LicenseStatus GetLicenseStatus()
        {
            return licenseStatus;
        }

        public string GetUUID()
        {
            return UUID;
        }

        private LicenseTypes ParseLicTypes(string p_string)
        {
            return p_string switch
            {
                "Annual subscription" => LicenseTypes.ANNUAL_SUBSCRIPTION,
                "Trial" => LicenseTypes.TRIAL,
                "Development" => LicenseTypes.DEVELOPMENT,
                _ => LicenseTypes.UNKNOWN,
            };
        }

        private LicenseStatus ParseLicStatus(string p_string)
        {
            return p_string switch
            {
                "active" => LicenseStatus.ACTIVE,
                "inactive" => LicenseStatus.INACTIVE,
                _ => LicenseStatus.INACTIVE,
            };
        }
    }
}