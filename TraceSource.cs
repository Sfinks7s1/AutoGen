using System;
using System.IO;
using System.Text;

namespace Auto
{
    internal static class TraceSource
    {
        private static readonly object Sync = new object();

        public static void CreateTraceFile()
        {
            if (!Directory.Exists(@"C:\Temp\"))
            {
                Directory.CreateDirectory(@"C:\Temp\");
            }

            if (File.Exists(LogFileFullPath))
            {
                File.Delete(LogFileFullPath);
                TraceMessage(TraceType.Information, "Create trace file");
            }
            else
            {
                TraceMessage(TraceType.Information, "Create trace file");
            }
        }

        internal static string LogFileFullPath { get; set; } = @"C:\Temp\AutocadTdmsPlugin.log";

        public static void TraceMessage(TraceType type, string msg)
        {
            try
            {
                string[] lines =
                    {
                        String.Concat( DateTime.Now.Date.ToShortDateString(),
                                       " ",
                                       DateTime.Now.TimeOfDay.ToString(),
                                       " ",
                                       "[" + type + "]",
                                       " ",
                                       msg)
                    };

                lock (Sync)
                {
                    File.AppendAllLines(LogFileFullPath, lines, Encoding.GetEncoding("Windows-1251"));
                }
            }
            catch
            {
                // Перехватываем все и ничего не делаем
            }
        }
    }
}