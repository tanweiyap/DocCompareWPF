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
        private static readonly string serverAddress = "https://hopie.tech:3501/";

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

        [ProtoMember(7)]
        private DateTime expiryWaiveDate;

        [ProtoMember(8)]
        private bool waived = false;

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
                licenseType = LicenseTypes.FREE;
                licenseStatus = LicenseStatus.ACTIVE;
                expiryDate = DateTime.Today.AddDays(7);
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
            FREE,
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
            INUSE,
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
                    DateTime returnDate = DateTime.Parse((string)resp["DATE"], CultureInfo.GetCultureInfo("de-de"));

                    TimeSpan diff = returnDate.Subtract(expiryDate); // if local date is in the future, we know that the trial was installed
                    expiryDate = returnDate;

                    if (diff.TotalDays < -7)
                    {
                        licenseStatus = LicenseStatus.INACTIVE;
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

        public async Task<LicServerResponse> ExtendTrial()
        {
            int status = await CheckServerStatus(serverAddress + "status");

            expiryDate = expiryDate.AddDays(7);
            licenseStatus = LicenseStatus.ACTIVE;

            if (status == 0)
            {
                IDictionary<string, string> licDict = new Dictionary<string, string>
                    {
                        { "UUID", UUID},
                        { "DATE", expiryDate.ToString("d",CultureInfo.GetCultureInfo("de-de"))}
                    };

                var content = new FormUrlEncodedContent(licDict);

                HttpResponseMessage msg = await client.PostAsync(serverAddress + "extendtrial", content);
                string readMsg = msg.Content.ReadAsStringAsync().Result;
                try
                {
                    JObject resp = JsonConvert.DeserializeObject<JObject>(readMsg);
                    DateTime returnDate = DateTime.Parse((string)resp["DATE"], CultureInfo.GetCultureInfo("de-de"));

                    TimeSpan diff = returnDate.Subtract(expiryDate); // if local date is in the future, we know that the trial was installed
                    if (diff.TotalDays < 0)
                    {
                        expiryDate = returnDate;
                        licenseStatus = LicenseStatus.INACTIVE;
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

        public void ConvertTrialToFree()
        {
            if (licenseType == LicenseTypes.TRIAL)
            {
                licenseType = LicenseTypes.FREE;
            }

        }

        public async Task<LicServerResponse> ActivateLicense(string userEmail, string licKey)
        {
            int status = await CheckServerStatus(serverAddress + "status");
            if (status == 0)
            {
                try
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

                    if ((string)resp["InUse"] == "true")
                    {
                        return LicServerResponse.INUSE;
                    }

                    if ((string)resp["CorrectKey"] == "true")
                    {
                        licenseType = ParseLicTypes((string)resp["LicType"]);
                        licenseStatus = ParseLicStatus((string)resp["LicStatus"]);
                        expiryDate = DateTime.Parse((string)resp["Expires"], CultureInfo.GetCultureInfo("de-de"));
                        email = userEmail;
                        key = licKey;
                        waived = false;
                        return LicServerResponse.OKAY;
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

        public async Task<LicServerResponse> RenewLicense()
        {
            int status = await CheckServerStatus(serverAddress + "status");
            if (status == 0)
            {
                IDictionary<string, string> licDict = new Dictionary<string, string>
                    {
                        { "Email", email},
                        { "LicKey", key},
                        { "UUID", UUID }
                    };

                var content = new FormUrlEncodedContent(licDict);

                HttpResponseMessage msg = await client.PostAsync(serverAddress + "renew", content);
                string readMsg = msg.Content.ReadAsStringAsync().Result;
                try
                {
                    JObject resp = JsonConvert.DeserializeObject<JObject>(readMsg);

                    if ((string)resp["CorrectKey"] == "true")
                    {
                        licenseType = ParseLicTypes((string)resp["LicType"]);
                        licenseStatus = ParseLicStatus((string)resp["LicStatus"]);
                        expiryDate = DateTime.Parse((string)resp["Expires"], CultureInfo.GetCultureInfo("de-de"));
                        key = (string)resp["LicKey"]; // new License key
                        return LicServerResponse.OKAY;
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
                }
                catch (Exception ex)
                {
                    ErrorHandling.ReportException(ex);
                }
            }
            else
            {
                if (waived == false)
                {
                    expiryWaiveDate = expiryDate.AddDays(7);
                    waived = true;
                }
                return LicServerResponse.UNREACHABLE; // server offline
            }

            return LicServerResponse.INVALID;
        }

        public async Task<bool> RemoveLicense()
        {
            int status = await CheckServerStatus(serverAddress + "status");
            if (status == 0)
            {
                IDictionary<string, string> licDict = new Dictionary<string, string>
                    {
                        { "Email", email},
                        { "LicKey", key},
                        { "UUID", UUID }
                    };

                var content = new FormUrlEncodedContent(licDict);

                HttpResponseMessage msg = await client.PostAsync(serverAddress + "remove", content);
                string readMsg = msg.Content.ReadAsStringAsync().Result;
                try
                {
                    JObject resp = JsonConvert.DeserializeObject<JObject>(readMsg);

                    if ((string)resp["Status"] == "OK")
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandling.ReportException(ex);
                }
            }
            else
            {
                return false;
            }

            return false;
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

        public async Task<List<string>> CheckUpdate(string currVersion, string localeType)
        {
            int status = await CheckServerStatus(serverAddress + "status");
            if (status == 0)
            {
                IDictionary<string, string> licDict = new Dictionary<string, string>
                    {
                        { "EN_DE", localeType}
                    };

                var content = new FormUrlEncodedContent(licDict);

                HttpResponseMessage msg = await client.PostAsync(serverAddress + "app-release", content);
                string readMsg = msg.Content.ReadAsStringAsync().Result;
                try
                {
                    JObject resp = JsonConvert.DeserializeObject<JObject>(readMsg);

                    if ((string)resp["Version"] != currVersion)
                    {
                        List<string> res = new List<string>();
                        res.Add((string)resp["Version"]);
                        res.Add((string)resp["Link"]);
                        res.Add((string)resp["Info"]);
                        return res;
                    }
                    else
                    {
                        return null;
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandling.ReportException(ex);
                }
            }

            return null;
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
                LicenseTypes.FREE => "Free",
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

        public DateTime GetExpiryWaiveDate()
        {
            return expiryWaiveDate;
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
                "Free" => LicenseTypes.FREE,
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