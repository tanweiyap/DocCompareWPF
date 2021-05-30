using System;
using System.IO;
using System.Threading;

namespace LocalLogging
{
    public class LoggingClass
    {
        private static string workingDir = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), ".2compare", "Logs");
        private static ReaderWriterLockSlim readerWriterLock = new ReaderWriterLockSlim();

        public static void WriteLog(string msg)
        {
            if (Directory.Exists(workingDir) == false)
            {
                Directory.CreateDirectory(workingDir);
            }

            readerWriterLock.EnterWriteLock();
            try
            {
                using StreamWriter writer = new StreamWriter(Path.Join(workingDir, "LogFile_" + DateTime.Now.ToString("yy_MM_dd") + ".log"), true);

                writer.WriteLineAsync("[" + DateTime.Now.ToString("u") + "] " + msg);
            }
            finally
            {
                readerWriterLock.ExitWriteLock();
            }

        }

        public static void ClearLog()
        {
            if (Directory.Exists(workingDir) == false)
            {
                Directory.CreateDirectory(workingDir);
            }

            readerWriterLock.EnterWriteLock();
            try
            {
                using StreamWriter writer = new StreamWriter(Path.Join(workingDir, "LogFile_" + DateTime.Now.ToString("yy_MM_dd") + ".log"), false);
                writer.WriteLineAsync("----------------New Log File----------------");
            }
            finally
            {
                readerWriterLock.ExitWriteLock();
            }
        }

    }
}
