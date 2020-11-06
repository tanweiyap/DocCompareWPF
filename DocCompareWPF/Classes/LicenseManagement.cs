using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ProtoBuf;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Management;
using System.Net.Http;
using System.Threading.Tasks;

namespace DocCompareWPF.Classes
{
    [ProtoContract]
    internal class LicenseManagement
    {
        private static readonly HttpClient client = new HttpClient() { Timeout = new TimeSpan(0, 0, 20) };

        //private static readonly string LocalDirectory = Directory.GetCurrentDirectory();
        private static readonly string serverAddress = "http://licserver.portmap.host:46928/";

        [ProtoMember(1)]
        private LicenseTypes licenseType;

        [ProtoMember(2)]
        private LicenseStatus licenseStatus;

        [ProtoMember(3)]
        private DateTime expiryDate;

        [ProtoMember(4)]
        private string UUID;


        [ProtoMember(5)]
        private string email = "";


        [ProtoMember(6)]
        private string key = "";

        public LicenseManagement()
        {
            
        }

        public void Init()
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
                licenseType = LicenseTypes.TRIAL;
                licenseStatus = LicenseStatus.ACTIVE;
                expiryDate = DateTime.Today.AddDays(14);
                //expiryDate = DateTime.Today.Subtract(TimeSpan.FromDays(2));
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
            INVALID,
        };

        public async Task<LicServerResponse> ActivateTrial()
        {
            int status = await CheckServerStatus(serverAddress + "status");
            if (status == 0)
            {
                IDictionary<string, string> licDict = new Dictionary<string, string>
                    {
                        { "UUID", UUID},
                        { "DATE", expiryDate.ToString("d",CultureInfo.GetCultureInfo("de-de"))}
                    };

                var content = new FormUrlEncodedContent(licDict);

                HttpResponseMessage msg = await client.PostAsync(serverAddress + "trial", content);
                string readMsg = msg.Content.ReadAsStringAsync().Result;
                try
                {
                    JObject resp = JsonConvert.DeserializeObject<JObject>(readMsg);
                    DateTime returnDate = DateTime.Parse((string)resp["DATE"],CultureInfo.GetCultureInfo("de-de"));

                    TimeSpan diff = returnDate.Subtract(expiryDate); // if local date is in the future, we know that the trial was installed
                    if(diff.TotalDays < 0)
                    {
                        expiryDate = returnDate;
                        return LicServerResponse.INVALID;
                    }

                    return LicServerResponse.OKAY;
                }
                catch (Exception ex)
                {
                    ErrorHandling.ReportException(ex);
                }
            }
            else
            {
                return LicServerResponse.UNREACHABLE; // server offline
            }

            return LicServerResponse.INVALID;
        }

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
                try
                {
                    JObject resp = JsonConvert.DeserializeObject<JObject>(readMsg);

                    if ((string)resp["CorrectKey"] == "true")
                    {
                        licenseType = ParseLicTypes((string)resp["LicType"]);
                        licenseStatus = ParseLicStatus((string)resp["LicStatus"]);
                        expiryDate = DateTime.Parse((string)resp["Expires"],CultureInfo.GetCultureInfo("de-de"));
                        email = userEmail;
                        key = licKey;
                    }
                    else
                    {
                        licenseType = LicenseTypes.UNKNOWN;
                        licenseStatus = LicenseStatus.INACTIVE;
                        expiryDate = DateTime.Now;

                        if (ParseLicTypes((string)resp["LicType"]) == LicenseTypes.UNKNOWN)
                        {
                            return LicServerResponse.ACCOUNT_NOT_FOUND; //
                        }

                        return LicServerResponse.KEY_MISMATCH; // wrong key
                    }
                }catch(Exception ex)
                {
                    ErrorHandling.ReportException(ex);
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

        public string GetLicenseTypesString()
        {
            return licenseType switch
            {
                LicenseTypes.ANNUAL_SUBSCRIPTION => "Annual subscription",
                LicenseTypes.DEVELOPMENT => "Developer license",
                LicenseTypes.TRIAL => "Trial license",
                LicenseTypes.UNKNOWN => "Unknown license type",
                _ => "",
            };
        }

        public LicenseStatus GetLicenseStatus()
        {
            return licenseStatus;
        }

        public DateTime GetExpiryDate()
        {
            return expiryDate;
        }

        public string GetExpiryDateString()
        {
            return expiryDate.ToShortDateString();
        }

        public string GetUUID()
        {
            return UUID;
        }

        public string GetEmail()
        {
            return email;
        }

        public string GetKey()
        {
            return key;
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