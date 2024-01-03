using System;
using System.IO;

namespace Rich_Text_Processor
{
    public static class Logger
    {
        static Logger() => File.WriteAllText(LogFilePath, string.Empty); // Ensure the log file is created or cleared when the application starts.

        public static object LockObj { get; } = new object();

        public static string LogFilePath { get; } = "Rich_Text_Processor.log.rtf";

        public static void LogError(string message) => LogMessage($"ERROR: {message}");

        public static void LogInfo(string message) => LogMessage($"INFO: {message}");

        private static void LogMessage(string message)
        {
            lock (LockObj)
            {
                string formattedMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
                Console.WriteLine(formattedMessage);

                File.AppendAllText(LogFilePath, $@"\par {formattedMessage}\par\par"); // Append the log entry to the RTF file.
            }
        }
    }
}