using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

namespace DocCompareWPF.Classes
{
    internal class Error
    {
        public string Callstack;
        public string ErrMessage;
        public string ErrType;
        public string Time;
    }

    internal class ErrorHandling
    {
        private static readonly HttpClient client = new HttpClient();
        private static readonly string LocalDirectory = Directory.GetCurrentDirectory();
        private static readonly string serverAddress = "http://errorlogging.portmap.host:44200/posterror";
        static public void ReportError(string ErrType, string CallStack, string Message)
        {
            Error err = new Error
            {
                ErrType = ErrType,
                Callstack = CallStack,
                ErrMessage = Message,
                Time = DateTime.Now.ToString(),
            };

            try
            {
                SendLog(err);
                //WriteLog(err);
            }
            catch
            {
            }
        }

        static public void ReportException(Exception ex)
        {
            Error err = new Error
            {
                ErrType = ex.GetType().ToString(),
                Callstack = ex.StackTrace,
                ErrMessage = ex.Message,
                Time = DateTime.Now.ToString(),
            };

            try
            {
                SendLog(err);
                //WriteLog(err);
            }
            catch
            {
            }
        }

        static private async Task SendLog(Error err)
        {
            IDictionary<string, string> errDict = new Dictionary<string, string>
            {
                { "ErrType", err.ErrType },
                { "CallStack", err.Callstack },
                { "ErrMessage", err.ErrMessage },
                { "Time", err.Time }
            };

            var content = new FormUrlEncodedContent(errDict);

            _ = await client.PostAsync(serverAddress, content);
        }

        /*
        static private void WriteLog(Error err)
        {
        }
        */
    }
}