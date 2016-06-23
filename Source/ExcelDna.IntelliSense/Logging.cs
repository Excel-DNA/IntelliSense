﻿//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Diagnostics;
using System.Globalization;
using System.Security;

namespace ExcelDna.IntelliSense
{
    // This class supports internal logging, implemented with the System.Diagnostics tracing implementation.

    // Add a trace listener for the ExcelDna.Integration source which logs warnings and errors to the LogDisplay 
    // (only popping up the window for errors).
    // Verbose logging can be configured via the .config file

    // We define a TraceSource called ExcelDna.Integration (that is also exported to ExcelDna.Loader and called from there)
    // We consolidate the two assemblies against a single TraceSource, since ExcelDna.Integration is the only public contract,
    // and we expect to move more of the registration into the ExcelDna.Integration assembly in future.

    // DOCUMENT: Info on custom TraceSources etc: https://msdn.microsoft.com/en-us/magazine/cc300790.aspx
    //           and http://blogs.msdn.com/b/kcwalina/archive/2005/09/20/tracingapis.aspx

    #region Microsoft License
    // The logging helper implementation here is adapted from the Logging.cs file for System.Net
    // Taken from https://github.com/Microsoft/referencesource/blob/c697a4b9782dc8c85c02344a847cb68163702aa7/System/net/System/Net/Logging.cs
    // Under the following license:
    //
    // The MIT License (MIT)

    // Copyright (c) Microsoft Corporation

    // Permission is hereby granted, free of charge, to any person obtaining a copy 
    // of this software and associated documentation files (the "Software"), to deal 
    // in the Software without restriction, including without limitation the rights 
    // to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
    // copies of the Software, and to permit persons to whom the Software is 
    // furnished to do so, subject to the following conditions: 

    // The above copyright notice and this permission notice shall be included in all 
    // copies or substantial portions of the Software. 

    // THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
    // IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
    // FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
    // AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
    // LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
    // OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE 
    // SOFTWARE.
    #endregion

    // NOTE: To simplify configuration (so that we provide one TraceSource per referenced assembly) and still allow some grouping
    //       we use the EventId to define a trace event classification.
    // These names are not surfaced to the Trace Listeners - just the IDs, so we should document them.
    enum IntelliSenseTraceEventId
    {
        Initialization = 1,
        Monitor = 2,
        WinEvents = 3,
        WindowWatcher = 4,
        Provider = 5,
        Display = 6,
    }

    // TraceLogger manages the IntelliSenseTraceSource that we use for logging.
    // It deals with lifetime (particularly closing the TraceSource if the library is unloaded).
    // The default configuration of the TraceSource is set here, and can be overridden in the .xll.config file.
    class TraceLogger
    {
        static volatile bool s_LoggingEnabled = true;
        static volatile bool s_LoggingInitialized;
        static volatile bool s_AppDomainShutdown;
        const string TraceSourceName = "ExcelDna.IntelliSense";
        internal static TraceSource IntelliSenseTraceSource;

        public static void Initialize()
        {
            if (!s_LoggingInitialized)
            {
                bool loggingEnabled = false;
                // DOCUMENT: By default the TraceSource is configured to source only Warning, Error and Fatal.
                //           the configuration can override this.
                IntelliSenseTraceSource = new TraceSource(TraceSourceName, SourceLevels.Warning);

                try
                {
                    loggingEnabled = (IntelliSenseTraceSource.Switch.ShouldTrace(TraceEventType.Critical));
                }
                catch (SecurityException)
                {
                    // These may throw if the caller does not have permission to hook up trace listeners.
                    // We treat this case as though logging were disabled.
                    Close();
                    loggingEnabled = false;
                }
                catch (Exception ex)
                {
                    // CONSIDER: What to do here - e.g. Configuration errors
                    Debug.Print($"ExcelDna.IntelliSense.TraceLogger - Error in Initialize: {ex.Message}");
                    Close();
                    loggingEnabled = false;
                }
                if (loggingEnabled)
                {
                    AppDomain currentDomain = AppDomain.CurrentDomain;
                    //currentDomain.UnhandledException += UnhandledExceptionHandler;
                    currentDomain.DomainUnload += AppDomainUnloadEvent;
                    currentDomain.ProcessExit += ProcessExitEvent;
                }
                s_LoggingEnabled = loggingEnabled;
                s_LoggingInitialized = true;
            }
        }

        static bool ValidateSettings(TraceSource traceSource, TraceEventType traceLevel)
        {
            if (!s_LoggingEnabled)
            {
                return false;
            }
            if (!s_LoggingInitialized)
            {
                Initialize();
            }
            if (traceSource == null || !traceSource.Switch.ShouldTrace(traceLevel))
            {
                return false;
            }
            if (s_AppDomainShutdown)
            {
                return false;
            }
            return true;
        }

        static void ProcessExitEvent(object sender, EventArgs e)
        {
            Close();
            s_AppDomainShutdown = true;
        }

        static void AppDomainUnloadEvent(object sender, EventArgs e)
        {
            Close();
            s_AppDomainShutdown = true;
        }

        static void Close()
        {
            if (IntelliSenseTraceSource != null)
                IntelliSenseTraceSource.Close();
        }

    }

    class Logger
    {
        int _eventId;

        Logger(IntelliSenseTraceEventId traceEventId)
        {
            _eventId = (int)traceEventId;
        }

        void Log(TraceEventType eventType, string message)
        {
            try
            {
                TraceLogger.IntelliSenseTraceSource.TraceEvent(eventType, _eventId, message);
            }
            catch (Exception e)
            {
                Debug.Print("ExcelDna.IntelliSense - Logger.Log error: " + e.Message);
            }
        }

        void Log(TraceEventType eventType, string format, params object[] args)
        {
            try
            {
                TraceLogger.IntelliSenseTraceSource.TraceEvent(eventType, _eventId, format, args);
            }
            catch (Exception e)
            {
                Debug.Print("ExcelDna.IntelliSense - Logger.Log error: " + e.Message);
            }
        }

        public void Verbose(string message)
        {
            Log(TraceEventType.Verbose, message);
        }

        public void Verbose(string format, params object[] args)
        {
            Log(TraceEventType.Verbose, format, args);
        }

        public void Info(string message)
        {
            Log(TraceEventType.Information, message);
        }

        public void Info(string format, params object[] args)
        {
            Log(TraceEventType.Information, format, args);
        }

        public void Warn(string message)
        {
            Log(TraceEventType.Warning, message);
        }

        public void Warn(string format, params object[] args)
        {
            Log(TraceEventType.Warning, format, args);
        }

        public void Error(string message)
        {
            Log(TraceEventType.Error, message);
        }

        public void Error(string format, params object[] args)
        {
            Log(TraceEventType.Error, format, args);
        }

        public void Error(Exception ex, string message, params object[] args)
        {
            if (args != null)
            {
                try
                {
                    message = string.Format(CultureInfo.InvariantCulture, message, args);
                }
                catch (Exception fex)
                {
                    Debug.Print("ExcelDna.IntelliSense - Logger.Error formatting exception " + fex.Message);
                }
            }
            Log(TraceEventType.Error, "{0} : {1} - {2}", message, ex.GetType().Name, ex.Message);
        }

        static internal Logger Initialization { get; } = new Logger(IntelliSenseTraceEventId.Initialization);
        static internal Logger Provider { get; } = new Logger(IntelliSenseTraceEventId.Provider);
        static internal Logger WinEvents { get; } = new Logger(IntelliSenseTraceEventId.WinEvents);
        static internal Logger WindowWatcher { get; } = new Logger(IntelliSenseTraceEventId.WindowWatcher);
        static internal Logger Display { get; } = new Logger(IntelliSenseTraceEventId.Display);
        static internal Logger Monitor { get; } = new Logger(IntelliSenseTraceEventId.Monitor);
    }
}
