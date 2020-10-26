using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;

namespace DocCompareWPF.Classes
{
    internal class Error
    {
        public string Callstack;
        public string ErrMessage;
        public string ErrType;
    }

    internal class ErrorHandling
    {
        private static readonly string serverAddress = "http://errorlogging.portmap.host:44200/posterror";
        private static readonly string LocalDirectory = Directory.GetCurrentDirectory();
        private static readonly HttpClient client = new HttpClient();

        static public void ReportError(string ErrType, string CallStack, string Message)
        {
            Error err = new Error
            {
                ErrType = ErrType,
                Callstack = CallStack,
                ErrMessage = Message,
            };

            try
            {
                bool ret = SendLog(err);
                if (ret == false)
                    WriteLog(err);
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
            };
        }

        static private void WriteLog(Error err)
        {
        }

        static private bool SendLog(Error err)
        {
            bool ret = false;

            IDictionary<string, string> errDict = new Dictionary<string, string>
            {
                { "ErrType", err.ErrType },
                { "CallStack", err.Callstack },
                { "ErrMessage", err.ErrMessage }
            };

            var content = new FormUrlEncodedContent(errDict);

            var response = client.PostAsync(serverAddress, content);

            ret = true;

            return ret;
        }
    }
}