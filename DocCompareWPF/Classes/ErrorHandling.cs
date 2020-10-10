using System;

namespace DocCompareWPF.Classes
{
    internal class Error
    {
        public string ErrType;
        public string Callstack;
        public string ErrMessage;
    }

    internal class ErrorHandling
    {
        private string serverAddress = "tanweiyap@gmail.com";

        static public void ReportException(Exception ex)
        {
            Error err = new Error
            {
                ErrType = ex.GetType().ToString(),
                Callstack = ex.StackTrace,
                ErrMessage = ex.Message,
            };
        }

        static public void ReportError(string ErrType, string CallStack, string Message)
        {
            Error err = new Error
            {
                ErrType = ErrType,
                Callstack = CallStack,
                ErrMessage = Message,
            };
        }
    }
}