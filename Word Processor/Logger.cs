using System;
using System.IO;

namespace Rich_Text_Processor
{
    public enum LogLevel { Info, Warning, Error }

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

                string rtfFormatting; // Using different RTF formatting based on log level
                switch (level)
                {
                    case LogLevel.Info:
                        rtfFormatting = @"\cf0";
                        break;
                    case LogLevel.Warning:
                        rtfFormatting = @"\cf1";
                        break;
                    case LogLevel.Error:
                        rtfFormatting = @"\cf2";
                        break;
                    default:
                        rtfFormatting = @"\cf0"; // Default to INFO
                        break;
                }

                File.AppendAllText(LogFilePath, $@"\par {rtfFormatting} {formattedMessage}\par\par");
            }
        }
    }
}
