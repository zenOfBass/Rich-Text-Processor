using System;
using System.IO;

namespace Rich_Text_Processor
{
    public enum LogLevel
    {
        Info,
        Warning,
        Error
    }

    public static class Logger
    {
        static Logger() => File.WriteAllText(LogFilePath, string.Empty);

        private static readonly object LockObj = new object();

        public static string LogFilePath { get; private set; } = "Rich_Text_Processor.log.rtf";

        public static void SetLogFilePath(string path)
        {
            if (!string.IsNullOrEmpty(path)) LogFilePath = path;
        }

        public static void Log(LogLevel level, string message)
        {
            lock (LockObj)
            {
                string formattedMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {level}: {message}";
                Console.WriteLine(formattedMessage);

                // Using different RTF formatting based on log level
                string rtfFormatting = level switch
                {
                    LogLevel.Info => @"\cf0",
                    LogLevel.Warning => @"\cf1",
                    LogLevel.Error => @"\cf2",
                    _ => @"\cf0", // Default to INFO
                };

                File.AppendAllText(LogFilePath, $@"\par {rtfFormatting} {formattedMessage}\par\par");
            }
        }
    }
}
