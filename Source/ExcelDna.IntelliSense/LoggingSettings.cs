﻿using Microsoft.Win32;
using System;
using System.Diagnostics;

namespace ExcelDna.Logging
{
    internal class LoggingSettings
    {
        public SourceLevels SourceLevel { get; }
        public TraceEventType? DebuggerLevel { get; }
        public TraceEventType? FileLevel { get; }
        public string FileName { get; }

        public LoggingSettings()
        {
            SourceLevel = Enum.TryParse(GetCustomSetting("SOURCE_LEVEL", "SourceLevel"), out SourceLevels sourceLevelResult) ? sourceLevelResult : SourceLevels.Warning;

            if (Enum.TryParse(GetCustomSetting("DEBUGGER_LEVEL", "DebuggerLevel"), out TraceEventType debuggerLevelResult))
                DebuggerLevel = debuggerLevelResult;

            if (Enum.TryParse(GetCustomSetting("FILE_LEVEL", "FileLevel"), out TraceEventType fileLevelResult))
                FileLevel = fileLevelResult;

            FileName = GetCustomSetting("FILE_NAME", "FileName");
        }

        private static string GetCustomSetting(string environmentName, string registryName)
        {
            return Environment.GetEnvironmentVariable($"EXCELDNA_INTELLISENSE_DIAGNOSTICS_{environmentName}") ??
                (Registry.GetValue(@"HKEY_CURRENT_USER\Software\ExcelDna\IntelliSense\Diagnostics", registryName, null) as string ??
                Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\ExcelDna\IntelliSense\Diagnostics", registryName, null) as string);
        }
    }
}
