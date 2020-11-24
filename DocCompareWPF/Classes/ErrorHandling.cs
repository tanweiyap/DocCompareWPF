using System;
using System.Collections.Generic;
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
    internal class Status
    {
        public string Message;
        public string Type;
        public string Time;
    }

    internal class ErrorHandling
    {
        private static readonly HttpClient client = new HttpClient();
        //private static readonly string LocalDirectory = Directory.GetCurrentDirectory();
        private static readonly string serverAddressError = "http://18.157.228.39:3500/posterror";//"http://errorlogging.portmap.host:44200/posterror";
        private static readonly string serverAddressStatus = "http://18.157.228.39:3500/poststatus";//"http://errorlogging.portmap.host:44200/poststatus";
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
#pragma warning disable CS4014 // Da auf diesen Aufruf nicht gewartet wird, wird die Ausführung der aktuellen Methode vor Abschluss des Aufrufs fortgesetzt.
                SendLog(err);
#pragma warning restore CS4014 // Da auf diesen Aufruf nicht gewartet wird, wird die Ausführung der aktuellen Methode vor Abschluss des Aufrufs fortgesetzt.
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
#pragma warning disable CS4014 // Da auf diesen Aufruf nicht gewartet wird, wird die Ausführung der aktuellen Methode vor Abschluss des Aufrufs fortgesetzt.
                SendLog(err);
#pragma warning restore CS4014 // Da auf diesen Aufruf nicht gewartet wird, wird die Ausführung der aktuellen Methode vor Abschluss des Aufrufs fortgesetzt.
                //WriteLog(err);
            }
            catch
            {
            }
        }

        static public void ReportStatus(string StatusType, string Msg)
        {
            Status status = new Status
            {
                Type = StatusType,
                Message = Msg,
                Time = DateTime.Now.ToString(),
            };

            try
            {
#pragma warning disable CS4014 // Da auf diesen Aufruf nicht gewartet wird, wird die Ausführung der aktuellen Methode vor Abschluss des Aufrufs fortgesetzt.
                SendStatus(status);
#pragma warning restore CS4014 // Da auf diesen Aufruf nicht gewartet wird, wird die Ausführung der aktuellen Methode vor Abschluss des Aufrufs fortgesetzt.
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

            _ = await client.PostAsync(serverAddressError, content);
        }

        static private async Task SendStatus(Status status)
        {
            IDictionary<string, string> statusDict = new Dictionary<string, string>
            {
                { "Type", status.Type },
                { "Message", status.Message },
                { "Time", status.Time }
            };

            var content = new FormUrlEncodedContent(statusDict);

            _ = await client.PostAsync(serverAddressStatus, content);
        }

        /*
        static private void WriteLog(Error err)
        {
        }
        */
    }
}