using System;
using System.IO;

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
        private readonly string serverAddress = "tanweiyap@gmail.com";
        private readonly string LocalDirectory = Directory.GetCurrentDirectory();

        static public void ReportError(string ErrType, string CallStack, string Message)
        {
            Error err = new Error
            {
                ErrType = ErrType,
                Callstack = CallStack,
                ErrMessage = Message,
            };
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

        static private void 
            WriteLog(Error err)
        {

        }
    }
}